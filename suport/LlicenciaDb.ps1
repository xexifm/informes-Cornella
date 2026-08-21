#requires -Version 5.1
<#
.SYNOPSIS
  La BASE DE DADES de llicencies: memoria de cada activitat entre informes.

.DESCRIPTION
  Un informe de llicencia gairebe mai va sol: primer el requeriment, despres el
  favorable pre-llicencia, despres el post. Fins ara la tria nomes es recordava
  mentre el programa estava obert ($st.MemAbans), o sigui que el segon informe
  de la mateixa llicencia tornava a demanar-ho TOT -inclosos els Id Firmadoc i
  els expedients, que ja s'havien escrit-.

  Aixo ho arregla igual que ho fan les activitats extraordinaries
  (Load-ActExtrRegistry a ActExtr.ps1): un JSON local, per ID GIA, amb tot el
  que l'assistent necessita per tornar-se a omplir sol.

  El fitxer viu a local\base-dades-llicencies\llicencies-db.json. 'local' es
  ignorada sencera pel git: la base porta noms de titulars i numeros
  d'expedient i NO ha de pujar mai al repositori public.

.NOTES
  Es carrega via dot-source des de Motor.ps1 (tambe en mode headless de
  proves). Les funcions de dades son PURES i es proven a Linux sense Word:
  nomes Show-LlicenciaDb toca WinForms.

  CONVENCIO ASCII: el codi no porta accents (PowerShell 5.1 els corromp segons
  amb quin encoding llegeixi el fitxer). El text accentuat que va a la pantalla
  es fa amb [char]0xNN.
#>

# ----------------------------------------------------------------------------
# On viu (per defecte local\base-dades-llicencies\, dins del clone pero fora
# del repositori). Es pot sobreescriure des de config.ps1 o des dels tests.
# ----------------------------------------------------------------------------
if (-not $Script:LlicDbDir) {
    $Script:LlicDbDir = if ($RepoRoot) { Get-LocalSubdir $RepoRoot 'Llicencies' } else { 'base-dades-llicencies' }
}
$Script:LlicDbFile = 'llicencies-db.json'

function Get-LlicenciaDbPath {
    return [string](Join-Path $Script:LlicDbDir $Script:LlicDbFile)
}

function New-LlicenciaDb {
    return [pscustomobject]@{ Version = 1; Llicencies = @() }
}

function Load-LlicenciaDb {
    $path = Get-LlicenciaDbPath
    if (-not (Test-Path -LiteralPath $path)) { return (New-LlicenciaDb) }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return (New-LlicenciaDb) }
        $obj = $raw | ConvertFrom-Json
        if ($null -eq $obj.Llicencies) {
            Add-Member -InputObject $obj -NotePropertyName Llicencies -NotePropertyValue @() -Force
        }
        # ConvertFrom-Json desempaqueta els arrays d'un sol element.
        $obj.Llicencies = @($obj.Llicencies)
        return $obj
    } catch {
        # Una base il-legible no pot impedir fer l'informe: es comenca de zero.
        return (New-LlicenciaDb)
    }
}

function Save-LlicenciaDb($db) {
    if (-not (Test-Path -LiteralPath $Script:LlicDbDir)) {
        New-Item -ItemType Directory -Path $Script:LlicDbDir -Force | Out-Null
    }
    $db.Version = 1
    ($db | ConvertTo-Json -Depth 20) | Set-Content -LiteralPath (Get-LlicenciaDbPath) -Encoding UTF8
}

# ----------------------------------------------------------------------------
# Consulta i modificacio (PURES: nomes toquen l'objecte en memoria)
# ----------------------------------------------------------------------------
function Get-LlicenciaRecord($db, [string]$idGia) {
    if ($null -eq $db -or $null -eq $db.Llicencies) { return $null }
    $id = ([string]$idGia).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { return $null }
    foreach ($r in @($db.Llicencies)) {
        if (([string]$r.IdGia).Trim() -eq $id) { return $r }
    }
    return $null
}

# Insereix o SUBSTITUEIX per ID GIA (no en pot haver dues del mateix GIA).
function Set-LlicenciaRecord($db, $record) {
    if ($null -eq $db.Llicencies) {
        Add-Member -InputObject $db -NotePropertyName Llicencies -NotePropertyValue @() -Force
    }
    $id = ([string]$record.IdGia).Trim()
    $out = New-Object System.Collections.ArrayList
    $trobat = $false
    foreach ($r in @($db.Llicencies)) {
        if (([string]$r.IdGia).Trim() -eq $id) { [void]$out.Add($record); $trobat = $true }
        else { [void]$out.Add($r) }
    }
    if (-not $trobat) { [void]$out.Add($record) }
    $db.Llicencies = $out.ToArray()
    return $db
}

function Remove-LlicenciaRecord($db, [string]$idGia) {
    if ($null -eq $db -or $null -eq $db.Llicencies) { return $db }
    $id = ([string]$idGia).Trim()
    $out = New-Object System.Collections.ArrayList
    foreach ($r in @($db.Llicencies)) {
        if (([string]$r.IdGia).Trim() -eq $id) { continue }
        [void]$out.Add($r)
    }
    $db.Llicencies = $out.ToArray()
    return $db
}

# ----------------------------------------------------------------------------
# $st de l'assistent <-> fitxa de la base (totes dues PURES)
# ----------------------------------------------------------------------------
# Un mapa qualsevol (hashtable o PSCustomObject sortit del JSON) a hashtable.
# ConvertFrom-Json torna PSCustomObjects i l'assistent treballa amb hashtables;
# sense aixo, la primera cosa que es llegiria de la base petaria.
function _LlicDbAMapa($o) {
    $h = @{}
    if ($null -eq $o) { return $h }
    if ($o -is [System.Collections.IDictionary]) {
        foreach ($k in @($o.Keys)) { $h[[string]$k] = $o[$k] }
        return $h
    }
    foreach ($p in @($o.PSObject.Properties)) { $h[[string]$p.Name] = $p.Value }
    return $h
}

# La memoria d'un bloc de documentacio (ABANS o DESPRES) a una forma DESABLE.
#
# Els sub-punts es desen amb la CLAU EN TEXT ("0", "1"...): les claus d'un
# hashtable numeric no sobreviuen el pas per JSON, i tornarien com a text de
# totes maneres.
function ConvertTo-LlicenciaMemoria($mem) {
    $out = [ordered]@{}
    if ($null -eq $mem) { return $out }
    foreach ($k in @((_LlicDbAMapa $mem).Keys)) {
        $e = _LlicDbAMapa ((_LlicDbAMapa $mem)[$k])
        $vals = [ordered]@{}
        foreach ($n in @((_LlicDbAMapa $e['Valors']).Keys)) {
            $vals[[string]$n] = [string](_LlicDbAMapa $e['Valors'])[$n]
        }
        $subs = [ordered]@{}
        foreach ($n in @((_LlicDbAMapa $e['Subs']).Keys)) {
            $subs[[string]$n] = [bool](_LlicDbAMapa $e['Subs'])[$n]
        }
        $out[[string]$k] = [ordered]@{
            Marcat = [bool]$e['Marcat']
            Estat  = [string]$e['Estat']
            Valors = $vals
            Subs   = $subs
        }
    }
    return $out
}

# ...i de tornada: el que llegeix Select-LlicDocumentacio ($preSel). Les claus
# dels sub-punts tornen a ser NUMERIQUES, que es com les indexa la pantalla.
function ConvertFrom-LlicenciaMemoria($mem) {
    $out = @{}
    if ($null -eq $mem) { return $out }
    $m = _LlicDbAMapa $mem
    foreach ($k in @($m.Keys)) {
        $e = _LlicDbAMapa $m[$k]
        $vals = @{}
        $mv = _LlicDbAMapa $e['Valors']
        foreach ($n in @($mv.Keys)) { $vals[[string]$n] = [string]$mv[$n] }
        $subs = @{}
        $ms = _LlicDbAMapa $e['Subs']
        foreach ($n in @($ms.Keys)) {
            $i = 0
            if ([int]::TryParse([string]$n, [ref]$i)) { $subs[$i] = [bool]$ms[$n] }
        }
        $out[[string]$k] = @{ Marcat = [bool]$e['Marcat']; Estat = [string]$e['Estat']; Valors = $vals; Subs = $subs }
    }
    return $out
}

# L'estat de l'assistent -> la fitxa que es desa. PURA.
#
# $historial: les entrades que ja hi havia (per anar-hi afegint els informes
# generats sense perdre els anteriors).
function ConvertTo-LlicenciaRecord($st, $historial = @()) {
    $h = _LlicDbAMapa $st
    $header = _LlicDbAMapa $h['Header']
    $hist = New-Object System.Collections.ArrayList
    foreach ($x in @($historial)) { [void]$hist.Add($x) }
    return [ordered]@{
        IdGia         = [string]$header['ID_GIA']
        Actualitzat   = (Get-Date).ToString('o')
        Titular       = [string]$header['TITULAR']
        Adreca        = [string]$header['ADRECA']
        Activitat     = [string]$header['ACTIVITAT']
        Fase          = [string]$h['Fase']
        EsProvisional = [bool]$h['Prov']
        Classificacio = [string]$header['CLASSIFICACIO']
        Header        = (_LlicDbAMapa $h['Header'])
        Abans         = (ConvertTo-LlicenciaMemoria $h['MemAbans'])
        Despres       = (ConvertTo-LlicenciaMemoria $h['MemDespres'])
        ProjKeys      = @($h['ProjKeys'])
        ProjVals      = (_LlicDbAMapa $h['ProjVals'])
        Tecnic        = (_LlicDbAMapa $h['Tecnic'])
        Condicions    = [string]$h['Condicions']
        Historial     = $hist.ToArray()
    }
}

# ...i de tornada: omple $st amb el que consti a la fitxa. PURA (modifica el
# hashtable que se li passa i el retorna).
#
# La CAPCALERA no es toca: TITULAR, ADRECA i ACTIVITAT ja les omple l'Excel per
# ID GIA a Get-HeaderData, i el que hi ha a la base pot ser mes vell.
function Restore-LlicenciaState($record, $st) {
    if ($null -eq $record) { return $st }
    $r = _LlicDbAMapa $record
    $st['MemAbans']   = ConvertFrom-LlicenciaMemoria $r['Abans']
    $st['MemDespres'] = ConvertFrom-LlicenciaMemoria $r['Despres']
    $st['ProjKeys']   = @($r['ProjKeys'])
    $st['ProjVals']   = _LlicDbAMapa $r['ProjVals']
    $st['Tecnic']     = _LlicDbAMapa $r['Tecnic']
    $st['Condicions'] = [string]$r['Condicions']
    return $st
}

# Una linia d'historial. PURA.
function New-LlicenciaHistorial([string]$fase, [string]$fitxer) {
    return [ordered]@{
        Data   = (Get-Date).ToString('o')
        Fase   = $fase
        Fitxer = $fitxer
    }
}

# La data d'actualitzacio d'una fitxa, en format llegible. PURA.
function Get-LlicenciaDataText($record) {
    if ($null -eq $record) { return '' }
    $r = _LlicDbAMapa $record
    $d = [datetime]::MinValue
    if ([datetime]::TryParse([string]$r['Actualitzat'], [ref]$d)) { return $d.ToString('dd/MM/yyyy') }
    return ''
}

# El resum d'una fitxa per a la llista. PURA.
function Get-LlicenciaResum($record) {
    $r = _LlicDbAMapa $record
    $nAb = @((_LlicDbAMapa $r['Abans']).Keys | Where-Object { [bool](_LlicDbAMapa ((_LlicDbAMapa $r['Abans'])[$_]))['Marcat'] }).Count
    $nDe = @((_LlicDbAMapa $r['Despres']).Keys | Where-Object { [bool](_LlicDbAMapa ((_LlicDbAMapa $r['Despres'])[$_]))['Marcat'] }).Count
    return [pscustomobject]@{
        IdGia     = [string]$r['IdGia']
        Titular   = [string]$r['Titular']
        Fase      = [string]$r['Fase']
        Data      = (Get-LlicenciaDataText $record)
        Punts     = ($nAb + $nDe)
        Informes  = @($r['Historial']).Count
    }
}

# ----------------------------------------------------------------------------
# PANTALLA DE CONSULTA (l'unica part que toca WinForms)
# ----------------------------------------------------------------------------
# Llistat de llicencies a l'esquerra i, a la dreta, el que se n'ha desat: els
# punts marcats de cada bloc amb les seves dades (Id Firmadoc, expedients...).
# Les dades son EDITABLES aqui mateix: quan una s'ha entrat malament, corregir-la
# aqui evita haver de refer l'informe per arreglar-la.
#
# S'hi entra pel xip "Dades" de la fila de Llicencia del menu principal.
function Show-LlicenciaDb {
    $db = Load-LlicenciaDb

    $form = _NewForm
    $form.Text = 'Base de dades de llic' + [char]0x00E8 + 'ncies'
    $form.ClientSize = New-Object System.Drawing.Size(1080, 660)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(860, 520)

    # Vegeu CLAUDE.md: les funcions de la pantalla van totes en un hashtable,
    # perque .GetNewClosure() nomes copia els locals i una closure que en cridi
    # una altra es quedaria amb $null.
    $fn = @{}
    $ui = @{ Busy = $false }

    # ---- Esquerra: la llista ----------------------------------------------
    $panEsq = New-Object System.Windows.Forms.Panel
    $panEsq.Location = New-Object System.Drawing.Point(14, 66)
    $panEsq.Size = New-Object System.Drawing.Size(430, 520)
    $panEsq.Anchor = 'Top,Bottom,Left'
    [void]$form.Controls.Add($panEsq)

    $graella = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $graella          # posa Dock='Fill': ha d'anar DINS d'un panell
    $graella.ReadOnly = $true
    [void]$panEsq.Controls.Add($graella)
    foreach ($c in @(
        @{ N = 'IdGia';    T = 'ID GIA';  W = 70 },
        @{ N = 'Titular';  T = 'Titular'; W = 170 },
        @{ N = 'Fase';     T = 'Fase';    W = 95 },
        @{ N = 'Data';     T = 'Data';    W = 80 })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = [string]$c.N
        $col.HeaderText = [string]$c.T
        $col.Width = [int]$c.W
        [void]$graella.Columns.Add($col)
    }

    # ---- Dreta: el detall --------------------------------------------------
    $panDret = New-Object System.Windows.Forms.Panel
    $panDret.Location = New-Object System.Drawing.Point(456, 66)
    $panDret.Size = New-Object System.Drawing.Size(610, 520)
    $panDret.Anchor = 'Top,Bottom,Left,Right'
    $panDret.AutoScroll = $true
    $panDret.BorderStyle = 'FixedSingle'
    $panDret.BackColor = [System.Drawing.Color]::White
    [void]$form.Controls.Add($panDret)

    # Fila visible -> ID GIA (l'ordre de la graella pot no ser el de la base).
    $mapa = New-Object System.Collections.ArrayList

    $fn.Omple = {
        $ui.Busy = $true
        try {
            $graella.Rows.Clear()
            $mapa.Clear()
            $ordenades = @(@($db.Llicencies) | Sort-Object -Property @{ Expression = { [string]$_.Actualitzat }; Descending = $true })
            foreach ($r in $ordenades) {
                $res = Get-LlicenciaResum $r
                [void]$mapa.Add([string]$res.IdGia)
                [void]$graella.Rows.Add(@($res.IdGia, $res.Titular, $res.Fase, $res.Data))
            }
        } finally { $ui.Busy = $false }
    }.GetNewClosure()

    # Les caselles de text que s'estan editant: control -> on va el valor.
    $edicions = New-Object System.Collections.ArrayList

    $fn.Pinta = {
        param($idGia)
        $panDret.Controls.Clear()
        $edicions.Clear()
        $rec = Get-LlicenciaRecord $db ([string]$idGia)
        if ($null -eq $rec) {
            # Mai en silenci: si no es troba la fitxa, que es vegi.
            $avis = New-Object System.Windows.Forms.Label
            $avis.Location = New-Object System.Drawing.Point(10, 10)
            $avis.AutoSize = $true
            $avis.ForeColor = [System.Drawing.Color]::DimGray
            $avis.Text = 'Tria una llic' + [char]0x00E8 + 'ncia de la llista.'
            [void]$panDret.Controls.Add($avis)
            return
        }
        $y = 10

        $lb = New-Object System.Windows.Forms.Label
        $lb.Location = New-Object System.Drawing.Point(10, $y)
        $lb.MaximumSize = New-Object System.Drawing.Size(560, 0)
        $lb.AutoSize = $true
        $lb.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $lb.Text = ('ID GIA ' + [string]$rec.IdGia + '  ' + [char]0x00B7 + '  ' + [string]$rec.Titular)
        [void]$panDret.Controls.Add($lb)
        $y += [Math]::Max(24, $lb.PreferredHeight + 4)

        $lb2 = New-Object System.Windows.Forms.Label
        $lb2.Location = New-Object System.Drawing.Point(10, $y)
        $lb2.MaximumSize = New-Object System.Drawing.Size(560, 0)
        $lb2.AutoSize = $true
        $lb2.ForeColor = [System.Drawing.Color]::DimGray
        $lb2.Text = ([string]$rec.Activitat + '  ' + [char]0x00B7 + '  ' + [string]$rec.Adreca +
                     "`n" + 'Fase: ' + [string]$rec.Fase + '   Actualitzat: ' + (Get-LlicenciaDataText $rec) +
                     '   Informes: ' + @($rec.Historial).Count)
        [void]$panDret.Controls.Add($lb2)
        $y += [Math]::Max(24, $lb2.PreferredHeight + 12)

        foreach ($bloc in @(
            @{ Prop = 'Abans';   Titol = 'Documentaci' + [char]0x00F3 + ' ABANS de la resoluci' + [char]0x00F3 },
            @{ Prop = 'Despres'; Titol = 'Documentaci' + [char]0x00F3 + ' DESPR' + [char]0x00C9 + 'S de la resoluci' + [char]0x00F3 })) {
            $mem = _LlicDbAMapa $rec.($bloc.Prop)
            $marcats = @(@($mem.Keys) | Where-Object { [bool](_LlicDbAMapa $mem[$_])['Marcat'] } | Sort-Object)
            if ($marcats.Count -eq 0) { continue }

            $lbB = New-Object System.Windows.Forms.Label
            $lbB.Location = New-Object System.Drawing.Point(10, $y)
            $lbB.AutoSize = $true
            $lbB.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            $lbB.Text = [string]$bloc.Titol
            [void]$panDret.Controls.Add($lbB)
            $y += 24

            foreach ($clau in $marcats) {
                $e = _LlicDbAMapa $mem[$clau]
                $lbP = New-Object System.Windows.Forms.Label
                $lbP.Location = New-Object System.Drawing.Point(20, $y)
                $lbP.MaximumSize = New-Object System.Drawing.Size(550, 0)
                $lbP.AutoSize = $true
                $estatTxt = if ([string]$e['Estat'] -eq 'si') { 'es disposa' } else { 'no es disposa' }
                $lbP.Text = ([string]$clau + '  (' + $estatTxt + ')')
                [void]$panDret.Controls.Add($lbP)
                $y += [Math]::Max(20, $lbP.PreferredHeight + 2)

                $vals = _LlicDbAMapa $e['Valors']
                foreach ($nom in @($vals.Keys | Sort-Object)) {
                    $lbN = New-Object System.Windows.Forms.Label
                    $lbN.Location = New-Object System.Drawing.Point(40, ($y + 3))
                    $lbN.Size = New-Object System.Drawing.Size(150, 20)
                    $lbN.Text = [string]$nom + ':'
                    [void]$panDret.Controls.Add($lbN)
                    $tb = New-Object System.Windows.Forms.TextBox
                    $tb.Location = New-Object System.Drawing.Point(195, $y)
                    $tb.Size = New-Object System.Drawing.Size(360, 22)
                    $tb.Text = [string]$vals[$nom]
                    [void]$panDret.Controls.Add($tb)
                    [void]$edicions.Add(@{ Ctrl = $tb; Bloc = [string]$bloc.Prop; Clau = [string]$clau; Nom = [string]$nom })
                    $y += 26
                }
                $y += 6
            }
            $y += 8
        }
        # La RESTA del que es recorda: sense aixo la fitxa semblava buida encara
        # que hi hagues mitja llicencia desada.
        $altres = New-Object System.Collections.ArrayList
        $nProj = @($rec.ProjKeys).Count
        if ($nProj -gt 0) { [void]$altres.Add('Punts del projecte triats: ' + $nProj) }
        $tec = _LlicDbAMapa $rec.Tecnic
        $bitsTec = New-Object System.Collections.ArrayList
        foreach ($k in @('Tecnic','NumCol','Collegi','Data')) {
            if (-not [string]::IsNullOrWhiteSpace([string]$tec[$k])) { [void]$bitsTec.Add($k + ': ' + [string]$tec[$k]) }
        }
        if ($bitsTec.Count -gt 0) { [void]$altres.Add('Tecnic redactor  ' + [char]0x00B7 + '  ' + ($bitsTec -join '   ')) }
        if (-not [string]::IsNullOrWhiteSpace([string]$rec.Condicions)) {
            $c = [string]$rec.Condicions
            if ($c.Length -gt 200) { $c = $c.Substring(0, 200) + [char]0x2026 }
            [void]$altres.Add('Condicions: ' + $c)
        }
        foreach ($h in @($rec.Historial)) {
            $hh = _LlicDbAMapa $h
            [void]$altres.Add('Informe generat: ' + [string]$hh['Fase'] + '  ' + [char]0x00B7 + '  ' +
                              (Split-Path -Leaf ([string]$hh['Fitxer'])))
        }
        if ($altres.Count -gt 0) {
            $lbA = New-Object System.Windows.Forms.Label
            $lbA.Location = New-Object System.Drawing.Point(10, $y)
            $lbA.AutoSize = $true
            $lbA.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            $lbA.Text = 'La resta de l' + [char]0x2019 + 'informe'
            [void]$panDret.Controls.Add($lbA)
            $y += 24
            foreach ($t in $altres) {
                $lbT = New-Object System.Windows.Forms.Label
                $lbT.Location = New-Object System.Drawing.Point(20, $y)
                $lbT.MaximumSize = New-Object System.Drawing.Size(550, 0)
                $lbT.AutoSize = $true
                $lbT.Text = [string]$t
                [void]$panDret.Controls.Add($lbT)
                $y += [Math]::Max(20, $lbT.PreferredHeight + 3)
            }
            $y += 8
        }
        if ($edicions.Count -eq 0 -and $altres.Count -eq 0) {
            $buit = New-Object System.Windows.Forms.Label
            $buit.Location = New-Object System.Drawing.Point(10, $y)
            $buit.AutoSize = $true
            $buit.ForeColor = [System.Drawing.Color]::DimGray
            $buit.Text = "D'aquesta llic" + [char]0x00E8 + 'ncia no se n' + [char]0x2019 + 'ha desat cap dada de documentaci' + [char]0x00F3 + '.'
            [void]$panDret.Controls.Add($buit)
        }
    }.GetNewClosure()

    # Escriu a la fitxa el que s'hagi canviat a les caselles i ho desa.
    $fn.Desa = {
        param($idGia)
        $rec = Get-LlicenciaRecord $db ([string]$idGia)
        if ($null -eq $rec) { return $false }
        $canvis = 0
        foreach ($ed in $edicions) {
            $mem = _LlicDbAMapa $rec.($ed.Bloc)
            if (-not $mem.ContainsKey($ed.Clau)) { continue }
            $e = _LlicDbAMapa $mem[$ed.Clau]
            $vals = $e['Valors']
            $nou = [string]$ed.Ctrl.Text
            if ($vals -is [System.Collections.IDictionary]) {
                if ([string]$vals[$ed.Nom] -ne $nou) { $vals[$ed.Nom] = $nou; $canvis++ }
            } else {
                if ([string]$vals.($ed.Nom) -ne $nou) {
                    Add-Member -InputObject $vals -NotePropertyName $ed.Nom -NotePropertyValue $nou -Force
                    $canvis++
                }
            }
        }
        if ($canvis -gt 0) { Save-LlicenciaDb $db }
        return $canvis
    }.GetNewClosure()

    # QUINA FILA HI HA TRIADA. Mira CurrentRow i, si no n'hi ha, les
    # seleccionades: posar '.Selected = $true' a ma NO mou el CurrentRow, i la
    # pantalla es quedava sense saber quina fila mirar.
    $fn.IdTriat = {
        $i = $graella.CurrentRow
        if ($null -eq $i -and @($graella.SelectedRows).Count -gt 0) { $i = @($graella.SelectedRows)[0] }
        if ($null -eq $i -or $i.Index -lt 0 -or $i.Index -ge $mapa.Count) { return '' }
        return [string]$mapa[$i.Index]
    }.GetNewClosure()

    $graella.add_SelectionChanged({
        if ($ui.Busy) { return }
        & $fn.Pinta (& $fn.IdTriat)
    }.GetNewClosure())
    # ...i tambe al clic: si la fila JA estava seleccionada, SelectionChanged no
    # es dispara i el detall no es tornava a pintar mai.
    $graella.add_CellClick({
        if ($ui.Busy) { return }
        & $fn.Pinta (& $fn.IdTriat)
    }.GetNewClosure())

    # La PRIMERA fila: cal moure-hi el CurrentCell (no nomes .Selected), si no
    # CurrentRow es queda a $null i el detall surt buit.
    $fn.TriaPrimera = {
        if ($graella.Rows.Count -le 0) { $panDret.Controls.Clear(); return }
        try { $graella.CurrentCell = $graella.Rows[0].Cells[0] } catch { }
        $graella.Rows[0].Selected = $true
        & $fn.Pinta (& $fn.IdTriat)
    }.GetNewClosure()

    & $fn.Omple
    & $fn.TriaPrimera
    # El detall es torna a pintar quan la finestra ja te handle: abans de
    # mostrar-la, el CurrentCell encara pot no estar posat.
    $form.add_Shown({ & $fn.TriaPrimera }.GetNewClosure())

    # ---- Botons ------------------------------------------------------------
    $btnTanca = New-Object System.Windows.Forms.Button
    $btnTanca.Text = 'Tancar'
    $btnTanca.Location = New-Object System.Drawing.Point(941, 606)
    $btnTanca.Size = New-Object System.Drawing.Size(125, 34)
    $btnTanca.Anchor = 'Bottom,Right'
    _StylePrimaryButton $btnTanca
    $btnTanca.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnTanca)

    $btnDesa = New-Object System.Windows.Forms.Button
    $btnDesa.Text = 'Desar els canvis'
    $btnDesa.Location = New-Object System.Drawing.Point(786, 606)
    $btnDesa.Size = New-Object System.Drawing.Size(145, 34)
    $btnDesa.Anchor = 'Bottom,Right'
    _StyleSecondaryButton $btnDesa
    $btnDesa.add_Click({
        $id = & $fn.IdTriat
        if ([string]::IsNullOrWhiteSpace($id)) { return }
        $n = & $fn.Desa $id
        $msg = if ($n -gt 0) { "S'han desat $n canvi(s)." } else { 'No hi havia res per canviar.' }
        [System.Windows.Forms.MessageBox]::Show($msg, 'Base de dades', 'OK', 'Information') | Out-Null
    }.GetNewClosure())
    [void]$form.Controls.Add($btnDesa)

    $btnEsb = New-Object System.Windows.Forms.Button
    $btnEsb.Text = 'Esborrar la fitxa'
    $btnEsb.Location = New-Object System.Drawing.Point(14, 606)
    $btnEsb.Size = New-Object System.Drawing.Size(150, 34)
    $btnEsb.Anchor = 'Bottom,Left'
    _StyleSecondaryButton $btnEsb
    $btnEsb.add_Click({
        $id = & $fn.IdTriat
        if ([string]::IsNullOrWhiteSpace($id)) { return }
        $r = [System.Windows.Forms.MessageBox]::Show(
            ("Segur que vols esborrar les dades desades de l'ID GIA " + $id + "?`n`n" +
             "Els informes ja generats NO es toquen; nomes es perd la memoria per al proper."),
            'Esborrar', 'YesNo', 'Warning')
        if ($r -ne 'Yes') { return }
        [void](Remove-LlicenciaRecord $db $id)
        Save-LlicenciaDb $db
        & $fn.Omple
        $panDret.Controls.Clear()
        & $fn.TriaPrimera
    }.GetNewClosure())
    [void]$form.Controls.Add($btnEsb)

    [void](_AddBrandHeader $form ('Base de dades de llic' + [char]0x00E8 + 'ncies') `
            ('El que es recorda de cada activitat per als informes seg' + [char]0x00FC + 'ents') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
}
