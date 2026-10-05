#requires -Version 5.1
<#
  Eina "Revisar requeriments": la finestra i la xarxa. Les decisions (vigencia,
  punts sense fitxa, les files de l'informe) son a RevisioDades.ps1; el
  recorregut dels enllacos, a Enllacos.ps1; la baixada de normativa, a
  Normativa.ps1 (Invoke-NormativaBaixada, la mateixa de l'eina Normativa).

  Fa quatre revisions (cada una es pot desmarcar) i en deixa un informe en Excel
  a local\revisions, amb una fila per cosa a fer:
    1. Punts de REQ1 sense fitxa d'informacio (la i).
    2. Enllacos que no responen, de tots els catalegs (text i fitxa).
    3. Normativa del cataleg que ja no es vigent, i la que la substitueix.
    4. (opcional) Baixar la normativa nova i les versions noves.
#>

# La vigencia d'UNA norma: @{ Estat; Detall; Substituta; SubstitutaId }. BOE:
# la pagina tal qual; Portal Juridic: el DOM que dibuixa l'Edge (l'etiqueta
# VIGENT es posa amb JavaScript). La resta ('?'): no hi ha manera fiable de
# saber-ho i val mes dir-ho.
function _RevVigenciaDe($e) {
    $u = [string]$e.Url
    try {
        if ((_NormativaFont $u) -eq 'boe') { return (_RevEstatBoe ([string](_NormativaGet $u).Content)) }
        if ($u -match '(?i)portaljuridic\.gencat\.cat') {
            # Primer SENSE navegador: les metadades ELI i la pagina del servidor.
            # L'Edge nomes si no ho diuen (al PC de l'usuari es penjava).
            $pj = _NormativaFontsPjur $u
            foreach ($t in @($pj.Textos)) {
                $v = _RevEstatEli ([string]$t)
                if ($v.Estat -ne '?') { return $v }
                $v = _RevEstatPjur ([string]$t)
                if ($v.Estat -ne '?') { return $v }
            }
            return (_RevEstatPjur (_NormativaDomEdge $u))
        }
    } catch {
        return @{ Estat = '?'; Detall = ('no s''ha pogut obrir: ' + $_.Exception.Message); Substituta = ''; SubstitutaId = '' }
    }
    return @{ Estat = 'n/a'; Detall = ''; Substituta = ''; SubstitutaId = '' }
}

function Get-RevisionsDir {
    if (-not $RepoRoot) { return '' }
    return [string](Get-LocalSubdir $RepoRoot 'Revisions')
}

function Invoke-RevisioRequeriments {
    $form = _NewForm
    $form.Text = 'Revisar requeriments'
    $form.ClientSize = New-Object System.Drawing.Size(720, 600)
    $form.MinimumSize = New-Object System.Drawing.Size(620, 500)
    $form.StartPosition = 'CenterScreen'
    $ui = @{ Corrent = $false; Cancel = $false; Informe = '' }

    $y = 70
    $chk = @{}
    foreach ($c in @(
        @{ K = 'Fitxes';    T = "Punts de REQ1 sense la fitxa d'informació (la i)"; On = $true },
        @{ K = 'Enllacos';  T = "Enllaços que no funcionen (tots els catàlegs)"; On = $true },
        @{ K = 'Vigencia';  T = "Normativa que ja no és vigent (BOE i Portal Jurídic)"; On = $true },
        @{ K = 'Baixa';     T = "Baixar la normativa nova i les versions noves (com l'eina Normativa)"; On = $false })) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = [string]$c.T
        $cb.Location = New-Object System.Drawing.Point(16, $y)
        $cb.AutoSize = $true
        $cb.Checked = [bool]$c.On
        $form.Controls.Add($cb)
        $chk[[string]$c.K] = $cb
        $y += 26
    }
    $lblNota = New-Object System.Windows.Forms.Label
    $lblNota.Text = "Les tres últimes necessiten Internet i poden trigar una estona. L'informe queda en un Excel a local\revisions."
    $lblNota.Location = New-Object System.Drawing.Point(16, ($y + 2))
    $lblNota.Size = New-Object System.Drawing.Size(688, 20)
    $lblNota.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblNota)
    $y += 30

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(16, $y)
    $bar.Size = New-Object System.Drawing.Size(688, 20)
    $bar.Anchor = 'Top,Left,Right'
    $form.Controls.Add($bar)
    $y += 28

    $log = New-Object System.Windows.Forms.TextBox
    $log.Multiline = $true; $log.ReadOnly = $true; $log.ScrollBars = 'Vertical'
    $log.BackColor = [System.Drawing.Color]::White
    $log.Font = New-Object System.Drawing.Font('Consolas', 9)
    $log.Location = New-Object System.Drawing.Point(16, $y)
    $log.Size = New-Object System.Drawing.Size(688, (548 - $y))
    $log.Anchor = 'Top,Bottom,Left,Right'
    $form.Controls.Add($log)

    $fn = @{}
    $fn.Log = { param($t) $log.AppendText($t + "`r`n"); [System.Windows.Forms.Application]::DoEvents() }.GetNewClosure()
    $fn.Pas = { $bar.Value = [Math]::Min($bar.Maximum, $bar.Value + 1); [System.Windows.Forms.Application]::DoEvents() }.GetNewClosure()
    $fn.Cancel = { [bool]$ui.Cancel }.GetNewClosure()

    $peu = _AddPeuBotons $form @(@{ Nom = 'Tanca'; Text = 'Tancar' }, @{ Nom = 'Informe'; Text = "Obrir l'informe" }) @(
        @{ Nom = 'Revisa'; Text = 'Revisar'; Estil = 'primari' }) 556 -Ancorat
    $btnTanca = $peu.Tanca; $btnRev = $peu.Revisa; $btnInf = $peu.Informe
    $btnInf.Enabled = $false
    $btnInf.add_Click({ if ($ui.Informe -and (Test-Path -LiteralPath $ui.Informe)) { try { Start-Process -FilePath $ui.Informe | Out-Null } catch { } } }.GetNewClosure())
    $btnTanca.add_Click({
        if ($ui.Corrent) { $ui.Cancel = $true; & $fn.Log "Aturant..."; return }
        $form.Close()
    }.GetNewClosure())
    $form.add_FormClosing({ param($s, $ev) if ($ui.Corrent) { $ui.Cancel = $true; $ev.Cancel = $true } }.GetNewClosure())

    $btnRev.add_Click({
        if ($ui.Corrent) { return }
        $ui.Corrent = $true; $ui.Cancel = $false
        $btnRev.Enabled = $false; $btnTanca.Text = 'Aturar'; $btnInf.Enabled = $false
        foreach ($c in $chk.Values) { $c.Enabled = $false }
        $log.Clear()
        $files = New-Object System.Collections.ArrayList
        $n = @{ Fitxes = 0; Enllacos = 0; Derogades = 0; Dubte = 0 }
        try {
            _NormativaPreparaXarxa
            $req1Json = Read-JsonFile (Join-Path $EstructuralsDir 'REQ1.json')

            # 1. PUNTS SENSE FITXA
            if ($chk.Fitxes.Checked) {
                & $fn.Log "1. Punts de REQ1 sense fitxa d'informació..."
                foreach ($p in @(_RevPuntsSenseFitxa $req1Json)) {
                    [void]$files.Add((_RevFila "Sense fitxa d'informació" 'REQ1' ([string]$p.Seccio + ' > ' + [string]$p.Punt) "El punt no té la fitxa d'informació (la i del Pas 3)." '' "Omple-la a l'editor de catàlegs (botó de la fitxa) o demana-la a Claude."))
                    $n.Fitxes++
                }
                & $fn.Log ('   ' + $n.Fitxes + ' punts sense fitxa.')
            }

            # 2. ENLLACOS
            if ($chk.Enllacos.Checked -and -not $ui.Cancel) {
                & $fn.Log "2. Enllaços de tots els catàlegs..."
                $cats = @(Get-ChildItem -LiteralPath $EstructuralsDir -Filter '*.json' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '0 *' } | Sort-Object Name)
                $perUrl = [ordered]@{}
                foreach ($f in $cats) {
                    foreach ($en in @(_EnllacosDeCataleg (Read-JsonFile $f.FullName))) {
                        $u = [string]$en.Url
                        if (-not $perUrl.Contains($u)) { $perUrl[$u] = New-Object System.Collections.ArrayList }
                        [void]$perUrl[$u].Add(@{ Cataleg = $f.BaseName; Punt = [string]$en.Punt; Camp = [string]$en.Camp })
                    }
                }
                $bar.Value = 0; $bar.Maximum = [Math]::Max(1, $perUrl.Count)
                foreach ($u in @($perUrl.Keys)) {
                    if ($ui.Cancel) { break }
                    $r = Test-EnllacViu $u
                    if (-not $r.Ok) {
                        $n.Enllacos++
                        & $fn.Log ('   NO RESPON [' + $r.Codi + '] ' + $u)
                        foreach ($on in @($perUrl[$u])) {
                            $camp = if ($on.Camp -eq 'fitxa') { "l'enllaç de la fitxa" } else { 'el text del punt' }
                            [void]$files.Add((_RevFila 'Enllaç trencat' $on.Cataleg $on.Punt ("No respon [" + $r.Codi + "] (" + $camp + ").") $u "Busca l'adreça nova i canvia-la a l'editor de catàlegs."))
                        }
                    }
                    & $fn.Pas
                }
                & $fn.Log ('   ' + $perUrl.Count + ' enllaços provats, ' + $n.Enllacos + ' que no responen.')
            }

            # 3. VIGENCIA
            if ($chk.Vigencia.Checked -and -not $ui.Cancel) {
                & $fn.Log "3. Vigència de la normativa..."
                # Una funcio, no $Script: aqui: som dins d'una closure.
                Reset-NormativaCaches
                $normes = @(Get-NormativaCataleg | Where-Object { -not $_.Guia -and -not $_.Colleccio -and -not $_.Derogada -and [string]$_.Url })
                $punts = @{}
                try { $punts = _NormativaPuntsReq1 $normes (Get-ParsedCataleg -path (Join-Path $EstructuralsDir 'REQ1.json')) } catch { }
                $bar.Value = 0; $bar.Maximum = [Math]::Max(1, $normes.Count)
                foreach ($e in $normes) {
                    if ($ui.Cancel) { break }
                    $v = _RevVigenciaDe $e
                    $id = [string]$e.Id
                    $pp = if ($punts.ContainsKey($id)) { (@($punts[$id]) -join '; ') } else { '' }
                    if ($v.Estat -eq 'derogada') {
                        $n.Derogades++
                        $subst = if ($v.Substituta) { ' Substituïda per: ' + $v.Substituta + $(if ($v.SubstitutaId) { ' (' + $v.SubstitutaId + ')' } else { '' }) + '.' } else { '' }
                        & $fn.Log ('   DEROGADA ' + $id + $subst)
                        $fer = if ($pp) { 'Revisar els punts de REQ1 que la citen i canviar-hi la norma; afegir la nova a la normativa (demana-ho a Claude).' } else { 'Treure-la de la normativa o marcar-la com a antiga (demana-ho a Claude).' }
                        [void]$files.Add((_RevFila 'Normativa derogada' 'REQ1' $(if ($pp) { $pp } else { '(no la cita cap punt)' }) ($id + ': ' + $v.Detall + $subst) ([string]$e.Url) $fer))
                    } elseif ($v.Estat -eq '?') {
                        $n.Dubte++
                        [void]$files.Add((_RevFila 'Vigència: mira-ho a mà' 'REQ1' $pp ($id + ": no s'ha pogut saber si és vigent. " + $v.Detall) ([string]$e.Url) "Obre l'enllaç i mira si diu que és vigent."))
                    }
                    & $fn.Pas
                }
                & $fn.Log ('   ' + $normes.Count + ' normes mirades: ' + $n.Derogades + ' derogades, ' + $n.Dubte + ' sense poder-ho saber.')
            }

            # 4. BAIXAR LA NORMATIVA NOVA
            if ($chk.Baixa.Checked -and -not $ui.Cancel) {
                & $fn.Log "4. Normativa nova i versions noves..."
                $totes = @(Get-NormativaCataleg)
                $dirN = Get-NormativaDir
                if (-not (Test-Path -LiteralPath $dirN)) { New-Item -ItemType Directory -Path $dirN -Force | Out-Null }
                $bar.Value = 0; $bar.Maximum = [Math]::Max(1, $totes.Count)
                $nb = Invoke-NormativaBaixada $totes $dirN $false $fn.Log $fn.Pas $fn.Cancel
                & $fn.Log (('   Noves: {0} · Actualitzades: {1} · Errors: {2}' -f $nb.Noves, $nb.Act, $nb.Err))
            }
        } catch {
            & $fn.Log ('ERROR: ' + $_.Exception.Message)
        } finally {
            if ($files.Count -eq 0) { [void]$files.Add((_RevFila 'Tot correcte' '' '' 'No hi ha res a fer.' '' '')) }
            try {
                $dirR = Get-RevisionsDir
                if (-not (Test-Path -LiteralPath $dirR)) { New-Item -ItemType Directory -Path $dirR -Force | Out-Null }
                $ui.Informe = Join-Path $dirR ('Revisio ' + (Get-Date).ToString('yyyy-MM-dd HHmm') + '.xlsx')
                [System.IO.File]::WriteAllBytes($ui.Informe, (_RevInformeXlsxBytes $files.ToArray()))
                & $fn.Log ''
                & $fn.Log ("Fet. L'informe: " + $ui.Informe)
                $btnInf.Enabled = $true
            } catch { & $fn.Log ("No s'ha pogut desar l'informe: " + $_.Exception.Message) }
            $ui.Corrent = $false
            $btnRev.Enabled = $true; $btnTanca.Text = 'Tancar'
            foreach ($c in $chk.Values) { $c.Enabled = $true }
        }
    }.GetNewClosure())

    [void](_AddBrandHeader $form 'Revisar requeriments' "Normativa vigent, enllaços i fitxes d'informació")
    [void]$form.ShowDialog()
    $form.Dispose()
}
