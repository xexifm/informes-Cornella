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

# EL JSON ES L'UNIC ORIGEN. NO hi ha cap copia dels textos escrita al codi.
#
# N'hi havia TRES -aqui, a docs\app.js (EMAIL_TEXTOS_DEFAULT) i al propi JSON- i
# el que les mantenia sincronitzades era un comentari. Van divergir: algu va
# afegir l'enllac castella de la seu des de la pantalla "Textos del correu" -que
# escriu el JSON- i les dues copies es van quedar enrere. O sigui que quan el
# mobil no podia llegir el fitxer enviava al titular una versio VELLA del correu,
# sense aquell enllac, i no ho deia ningu.
#
# Ara si el fitxer no hi es o no es valid, PETA amb un missatge clar. Es la
# mateixa regla que el projecte ja aplica als catalegs: val mes petar que enviar
# a un ciutada un text que no es el que toca. El fitxer esta al repositori, o
# sigui que 'git checkout -- docs/dades/email-textos.json' el recupera.

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

# Llegeix els textos del correu. Torna un ordered hashtable amb assumpte, cos i
# bcc. PETA si el fitxer no hi es, no es valid o li falta alguna clau: no hi ha
# cap valor de reserva a que caure (vegeu el comentari de mes amunt).
function _LoadEmailTextos {
    $path = _EmailTextosPath
    $o = Read-JsonFile $path
    if ($null -eq $o) {
        throw ("No s'han pogut llegir els textos del correu: " + $path + [Environment]::NewLine +
               "El fitxer no hi es o no es un JSON valid. Recupera'l amb:" + [Environment]::NewLine +
               "    git checkout -- docs/dades/email-textos.json")
    }
    $d = [ordered]@{}
    foreach ($k in @('assumpte', 'cos')) {
        if (-not $o.PSObject.Properties[$k] -or [string]::IsNullOrEmpty([string]$o.$k)) {
            throw ("Els textos del correu no tenen la clau '" + $k + "': " + $path)
        }
        $d[$k] = [string]$o.$k
    }
    $d['bcc'] = @(_EmailBccDeJson $o)
    return ,$d
}

# Llista de Copia Oculta (CCO) del JSON -> array d'objectes @{Addr; Default}.
#
# CONVENCIO: torna un array PLA i el crider hi posa @() (com
# _AutoFirmaCandidatePaths). NO fa 'return ,$out': combinar les dues coses hi
# posa la capa dues vegades i .Count val 1 -es la trampa que aquest projecte ja
# ha vist tres cops-.
# Les quatre adreces estaven escrites a EnviarCorreu.ps1 I a docs\app.js (i una
# tercera vegada a Recordatoris.ps1); ara son una clau mes del mateix fitxer, o
# sigui que canviar-ne una es tocar un sol lloc.
function _EmailBccDeJson($o) {
    $out = @()
    if ($null -eq $o -or -not $o.PSObject.Properties['bcc']) { return $out }
    foreach ($b in @($o.bcc)) {
        if ($null -eq $b) { continue }
        $addr = [string]$b.addr
        if ([string]::IsNullOrWhiteSpace($addr)) { continue }
        $out += @{ Addr = $addr.Trim(); Default = [bool]$b.def }
    }
    return $out
}

# Escriu el JSON (UTF-8 sense BOM). $obj = ordered hashtable amb les claus.
function _SaveEmailTextos($obj) {
    Write-JsonFile (_EmailTextosPath) $obj 5
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
    $btnDefault.Text = 'Descartar els canvis'
    $btnDefault.Location = New-Object System.Drawing.Point(136, 578)
    $btnDefault.Size = New-Object System.Drawing.Size(160, 30)
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

    # Torna a llegir el fitxer. Abans deia "Restaurar original" i posava una
    # copia dels textos escrita al codi; aquella copia ja no existeix -era una de
    # les tres que havien divergit- i el que de debo es vol des d'aqui es
    # desfer el que s'acaba d'escriure sense desar-ho.
    $btnDefault.add_Click({
        $r = [System.Windows.Forms.MessageBox]::Show('Vols tornar als textos que hi ha desats? Perdras els canvis que no hagis desat.', 'Textos del correu', 'YesNo', 'Question')
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        try {
            $desats = _LoadEmailTextos
            $tbA.Text = [string]$desats['assumpte']
            $tbC.Text = [string]$desats['cos'] -replace "`r?`n", "`r`n"
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Textos del correu', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    $btnSave.add_Click({
        if ([string]$tbC.Text -notmatch '\{REQUERIMENTS\}') {
            $r = [System.Windows.Forms.MessageBox]::Show("El cos no conte la variable {REQUERIMENTS}: els requeriments NO sortiran al correu.`n`nVols desar igualment?", 'Textos del correu', 'YesNo', 'Warning')
            if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        # Tornem a LF (\n) per al JSON (app.js parteix el cos per \n).
        # El bcc no s'edita des d'aqui, pero es del mateix fitxer: si no es
        # tornes a escriure, desar els textos se l'enduria.
        $out = [ordered]@{
            assumpte = [string]$tbA.Text
            cos      = ([string]$tbC.Text -replace "`r`n", "`n")
            bcc      = @(@($textos['bcc']) | ForEach-Object { [pscustomobject]@{ addr = [string]$_.Addr; def = [bool]$_.Default } })
        }
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
