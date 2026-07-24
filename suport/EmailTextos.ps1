#requires -Version 5.1
<#
.SYNOPSIS
  Editor dels TEXTOS del correu que l'app mòbil envia al titular (EmailJS).

.DESCRIPTION
  Els textos editables del correu (assumpte, frase d'introducció, bloc "Com
  presentar la documentació" en CA+ES i l'avís final en CA+ES) viuen a
  docs\dades\email-textos.json, que l'app mòbil (docs\app.js) llegeix. Aquesta
  eina els edita des del PC; en desar, s'escriu el JSON i es publica la propera
  vegada que es faci Actualitzar.bat (que ja puja docs\dades\*.json).

  Els textos admeten:
    - Placeholders: {ID_GIA} {ADRECA} {ACTIVITAT} {TITULAR} {DATA} (data d'avui).
    - **negreta** (a l'HTML del correu es passa a <b>; al text pla es treu).

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
    $d = [ordered]@{
        assumpte     = 'GIA {ID_GIA} Requeriments'
        introFrase   = "Aquestes s" + [char]0x00F3 + "n les defici" + [char]0x00E8 + "ncies que s'han detectat a la visita de l'activitat per part de l'Ajuntament el dia {DATA} i que s'han d'esmenar:"
        pujarTitolCa = "Com presentar la documentaci" + [char]0x00F3
        pujarTitolEs = "C" + [char]0x00F3 + "mo presentar la documentaci" + [char]0x00F3 + "n"
        pujarTextCa  = "Heu de presentar **tota la documentaci" + [char]0x00F3 + " alhora** (important: no la presenteu per parts), mitjan" + [char]0x00E7 + "ant una **inst" + [char]0x00E0 + "ncia gen" + [char]0x00E8 + "rica** de la seu electr" + [char]0x00F2 + "nica de l'Ajuntament de Cornell" + [char]0x00E0 + " de Llobregat:"
        pujarTextEs  = "Debe presentar **toda la documentaci" + [char]0x00F3 + "n a la vez** (importante: no la presente por partes), mediante una **instancia gen" + [char]0x00E9 + "rica** de la sede electr" + [char]0x00F3 + "nica del Ayuntamiento de Cornell" + [char]0x00E0 + " de Llobregat:"
        pujarUrl     = 'https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2'
        pujarInstrCa = "Indiqueu que la inst" + [char]0x00E0 + "ncia va **a l'atenci" + [char]0x00F3 + " del Departament d'Activitats** i feu-hi constar les dades de l'activitat:"
        pujarInstrEs = "Indique que la instancia va dirigida **a la atenci" + [char]0x00F3 + "n del Departamento de Actividades** y haga constar los datos de la actividad:"
        avisCa       = "IMPORTANT: aquest " + [char]0x00E9 + "s un correu autom" + [char]0x00E0 + "tic i no s'admeten respostes. Aquest llistat NO " + [char]0x00E9 + "s definitiu ni oficial i pot variar respecte del requeriment oficial que rebreu properament. Per a qualsevol consulta podeu adre" + [char]0x00E7 + "ar-vos al Departament d'Activitats de l'Ajuntament de Cornell" + [char]0x00E0 + " de Llobregat (Carrer de l'Energia, 97) o trucar al 93 377 02 12, extensi" + [char]0x00F3 + " 1227."
        avisEs       = "IMPORTANTE: este es un correo autom" + [char]0x00E1 + "tico y no se admiten respuestas. Este listado NO es definitivo ni oficial y puede variar respecto del requerimiento oficial que recibir" + [char]0x00E1 + " pr" + [char]0x00F3 + "ximamente. Para cualquier consulta puede dirigirse al Departamento de Actividades del Ayuntamiento de Cornell" + [char]0x00E0 + " de Llobregat (Calle de l'Energia, 97) o llamar al 93 377 02 12, extensi" + [char]0x00F3 + "n 1227."
    }
    return ,$d
}

# Metadades dels camps de l'editor (ordre, etiqueta, nombre de línies, ajuda).
function _EmailTextosFields {
    return @(
        @{ Key = 'assumpte';     Label = 'Assumpte del correu';                              Lines = 1; Hint = '{ID_GIA} = ID de l''activitat' }
        @{ Key = 'introFrase';   Label = 'Frase d''introduccio';                             Lines = 2; Hint = '{DATA} = data d''avui' }
        @{ Key = 'pujarTitolCa'; Label = 'Com presentar - Titol (catala)';                  Lines = 1; Hint = '' }
        @{ Key = 'pujarTitolEs'; Label = 'Com presentar - Titol (castella)';                Lines = 1; Hint = '' }
        @{ Key = 'pujarTextCa';  Label = 'Com presentar - Text (catala)';                   Lines = 3; Hint = '**negreta** amb dos asteriscs' }
        @{ Key = 'pujarTextEs';  Label = 'Com presentar - Text (castella)';                 Lines = 3; Hint = '' }
        @{ Key = 'pujarUrl';     Label = 'Com presentar - Enllac (seu electronica)';        Lines = 1; Hint = '' }
        @{ Key = 'pujarInstrCa'; Label = 'Com presentar - Instruccions (catala)';           Lines = 2; Hint = '' }
        @{ Key = 'pujarInstrEs'; Label = 'Com presentar - Instruccions (castella)';         Lines = 2; Hint = '' }
        @{ Key = 'avisCa';       Label = 'Avis final (catala)';                             Lines = 4; Hint = '' }
        @{ Key = 'avisEs';       Label = 'Avis final (castella)';                           Lines = 4; Hint = '' }
    )
}

# Llegeix el JSON (si hi es) fusionat sobre els valors per defecte. Retorna un
# ordered hashtable amb TOTES les claus.
function _LoadEmailTextos {
    $def = _DefaultEmailTextos
    $path = _EmailTextosPath
    if (Test-Path -LiteralPath $path) {
        try {
            $o = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
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
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $json = $obj | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

# ----------------------------------------------------------------------------
# INTERFICIE (WinForms) - nomes a Windows
# ----------------------------------------------------------------------------
function Invoke-EmailTextos {
    $textos = _LoadEmailTextos

    $form = _NewForm
    $form.Text = 'Textos del correu (mobil)'
    $form.ClientSize = New-Object System.Drawing.Size(760, 640)
    $form.MinimumSize = New-Object System.Drawing.Size(620, 480)

    # Panell desplaçable amb els camps (entre la capçalera i la barra de botons).
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(0, 56)
    $panel.Size = New-Object System.Drawing.Size(760, 528)
    $panel.Anchor = 'Top, Bottom, Left, Right'
    $panel.AutoScroll = $true
    [void]$form.Controls.Add($panel)

    $fBold = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $fHint = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Italic)
    $colHint = [System.Drawing.Color]::FromArgb(120, 128, 138)

    $boxes = @{}
    $y = 12
    foreach ($f in @(_EmailTextosFields)) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = [string]$f.Label
        $lbl.Font = $fBold
        $lbl.Location = New-Object System.Drawing.Point(16, $y)
        $lbl.AutoSize = $true
        [void]$panel.Controls.Add($lbl)
        $y += 20

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(16, $y)
        $lines = [int]$f.Lines
        $h = if ($lines -gt 1) { 20 * $lines + 6 } else { 24 }
        $tb.Size = New-Object System.Drawing.Size(710, $h)
        $tb.Anchor = 'Top, Left, Right'
        if ($lines -gt 1) { $tb.Multiline = $true; $tb.ScrollBars = 'Vertical'; $tb.WordWrap = $true }
        $tb.Text = [string]$textos[$f.Key]
        [void]$panel.Controls.Add($tb)
        $boxes[$f.Key] = $tb
        $y += $h + 3

        if (-not [string]::IsNullOrWhiteSpace([string]$f.Hint)) {
            $hint = New-Object System.Windows.Forms.Label
            $hint.Text = [string]$f.Hint
            $hint.Font = $fHint
            $hint.ForeColor = $colHint
            $hint.Location = New-Object System.Drawing.Point(18, $y)
            $hint.AutoSize = $true
            [void]$panel.Controls.Add($hint)
            $y += 16
        }
        $y += 8
    }

    # Barra inferior.
    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = 'Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(16, 598)
    $btnBack.Size = New-Object System.Drawing.Size(110, 30)
    $btnBack.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $btnBack
    [void]$form.Controls.Add($btnBack)

    $btnDefault = New-Object System.Windows.Forms.Button
    $btnDefault.Text = 'Restaurar originals'
    $btnDefault.Location = New-Object System.Drawing.Point(136, 598)
    $btnDefault.Size = New-Object System.Drawing.Size(160, 30)
    $btnDefault.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $btnDefault
    [void]$form.Controls.Add($btnDefault)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Desar'
    $btnSave.Location = New-Object System.Drawing.Point(624, 598)
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
        foreach ($k in @($def.Keys)) { if ($boxes.ContainsKey($k)) { $boxes[$k].Text = [string]$def[$k] } }
    }.GetNewClosure())

    $btnSave.add_Click({
        $out = [ordered]@{}
        foreach ($f in @(_EmailTextosFields)) { $out[$f.Key] = [string]$boxes[$f.Key].Text }
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
