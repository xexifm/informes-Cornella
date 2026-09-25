#requires -Version 5.1
<#
.SYNOPSIS
  Llicencia: les PANTALLES de l'assistent (WinForms, nomes Windows): la fase,
  la documentacio ABANS/DESPRES, el tecnic redactor i les condicions. Cada una
  torna un hashtable amb Nav ('fwd'/'back') i el que s'hi ha triat; la logica
  que es pot provar viu a LlicenciaDades.ps1.
  Vegeu la capcalera de LlicenciaDades.ps1 per al mapa del modul.
#>

# ----------------------------------------------------------------------------
# ASSISTENT (WinForms, nomes Windows)
# ----------------------------------------------------------------------------
# Pas 1: la FASE i si es llicencia provisional. Retorna @{ Nav; Fase; Prov }.
# $fases: quines fases s'ofereixen. Des que la Modificacio NO Substancial i el
# Traspas tenen entrada propia al menu, cada familia ensenya NOMES les seves:
# _LlicFases per a Llicencia i _MnsFases per a MNS/Traspas. Amb $null les
# ensenya totes (compatibilitat).
function Select-LlicFase($preFase, $preProv, $fases = $null, [string]$titol = '') {
    $llista = if ($null -ne $fases) { @($fases) } else { @(_LlicTotesLesFases) }
    $form = _NewForm
    $form.Text = if ($titol) { $titol } else { 'Llic' + [char]0x00E8 + 'ncia - Pas 1' }
    $form.ClientSize = New-Object System.Drawing.Size(520, 330)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 72)
    $lbl.Size = New-Object System.Drawing.Size(480, 20)
    $lbl.Text = 'Quin informe vols fer?'
    [void]$form.Controls.Add($lbl)

    $y = 98
    $radios = @{}
    foreach ($f in @($llista)) {
        $rb = New-Object System.Windows.Forms.RadioButton
        $rb.Location = New-Object System.Drawing.Point(30, $y)
        $rb.Size = New-Object System.Drawing.Size(460, 22)
        $rb.Text = [string]$f.Nom
        $rb.Checked = ([string]$preFase -eq [string]$f.Clau)
        [void]$form.Controls.Add($rb)
        $sub = New-Object System.Windows.Forms.Label
        $sub.Location = New-Object System.Drawing.Point(50, ($y + 21))
        $sub.Size = New-Object System.Drawing.Size(440, 18)
        $sub.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
        $sub.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $sub.Text = [string]$f.Sub
        [void]$form.Controls.Add($sub)
        $radios[[string]$f.Clau] = $rb
        $y += 46
    }
    if (-not ($radios.Values | Where-Object { $_.Checked })) {
        $perDefecte = _LlicFasePerDefecte $llista $preFase
        if ($radios.ContainsKey($perDefecte)) { $radios[$perDefecte].Checked = $true }
    }

    $cbProv = New-Object System.Windows.Forms.CheckBox
    $cbProv.Location = New-Object System.Drawing.Point(30, ($y + 6))
    $cbProv.AutoSize = $true
    $cbProv.Text = 'Llic' + [char]0x00E8 + 'ncia provisional'
    $cbProv.Checked = [bool]$preProv
    [void]$form.Controls.Add($cbProv)

    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Location = New-Object System.Drawing.Point(50, ($y + 28))
    $lbl2.Size = New-Object System.Drawing.Size(450, 32)
    $lbl2.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
    $lbl2.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lbl2.Text = ('Canvia el punt de compatibilitat (AMB en lloc d' + [char]0x2019 + 'Annex II) i, al requeriment, ' +
                  'hi afegeix l' + [char]0x2019 + 'ANNEX 1.')
    [void]$form.Controls.Add($lbl2)

    # La casella "Llicencia provisional" nomes te sentit a l'informe llarg: als
    # dos curts no canvia res del document, i deixar-la activa nomes despista.
    #
    # VA AQUI I NO MES AMUNT: .GetNewClosure() copia els VALORS del moment, o
    # sigui que una closure creada abans de $cbProv i $lbl2 se'ls quedaria a
    # $null (vegeu CLAUDE.md). I un clic en un radio dispara DOS esdeveniments
    # -el que es marca i el germa que es desmarca-, pero aqui es idempotent.
    $fnFase = @{}
    $fnFase.Refresca = {
        $curta = $false
        foreach ($k in @($radios.Keys)) { if ($radios[$k].Checked -and (_MnsEsFase $k)) { $curta = $true } }
        $cbProv.Enabled = (-not $curta)
        $lbl2.Visible = (-not $curta)
    }.GetNewClosure()
    foreach ($k in @($radios.Keys)) {
        $radios[$k].add_CheckedChanged({ & $fnFase.Refresca }.GetNewClosure())
    }
    & $fnFase.Refresca

    # ELS BOTONS, SOTA L'ULTIMA ETIQUETA. Estaven clavats a y=286 i la nota de
    # la llicencia provisional (y=264, alt 32) els trepitjava. Ara surten del
    # peu real de $lbl2, o sigui que si hi afegim una fase o una linia de text
    # baixen sols i la finestra creix amb ells.
    $yBotons = $lbl2.Bottom + 14
    $form.ClientSize = New-Object System.Drawing.Size(520, ($yBotons + 32 + 16))

    $res = @{ Nav = 'back' }
    $peu = _AddPeuBotons $form @(@{ Nom = 'Enrere'; Text = (_TxtEnrere) }) @(
        @{ Nom = 'Ok'; Text = 'Continuar'; Estil = 'primari' }) $yBotons
    $btnOk = $peu.Ok; $btnBack = $peu.Enrere
    $btnOk.add_Click({
        foreach ($k in $radios.Keys) { if ($radios[$k].Checked) { $res.Fase = $k } }
        $res.Prov = [bool]$cbProv.Checked
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())

    $btnBack.add_Click({ $form.Close() }.GetNewClosure())

    [void](_AddBrandHeader $form ('Llic' + [char]0x00E8 + 'ncia (Annex II / LL Prov)') 'Tria quin informe vols fer' 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

# Pas de tria de DOCUMENTACIO (blocs ABANS i DESPRES).
#
# ARBRE a l'esquerra + DETALL a la dreta, el mateix aspecte que el Pas 3
# (Select-Items, SeleccioItems.ps1): les seccions en negreta i els punts a
# dins. Abans era una llista plana amb 40 punts a la mateixa alcada, i abans
# encara una graella amb un boto "Omplir..." que obria un dialeg -que no
# s'assemblava a com s'omplen els camps a la resta del programa-.
#
# L'AGRUPACIO ES NOMES DE PANTALLA (_LlicAgrupaPunts): els punts es recorren en
# l'ordre del cataleg per muntar l'informe, o sigui que agrupar no en canvia
# l'ordre. Els punts sense seccio (els PROPIS, i els que es llegeixen d'un
# informe ja emes) van al primer nivell, sense capcalera.
#
# Al detall hi ha, segons el bloc:
#   - la tria "No es disposa / Es disposa" ($ambEstat);
#   - la frase del cataleg amb els [CAMP: ...] INLINE ($ambDades, nomes ABANS),
#     renderitzada amb _RenderRichInto (Camps.ps1) -la MATEIXA funcio que fa
#     servir REQ1-;
#   - les caselles dels SUB-PUNTS ($ambSubs, nomes DESPRES): els certificats
#     d'inscripcio i les inspeccions inicials no els te tothom.
#
# ELS CAMPS VAN PER PUNT, no al diccionari compartit: "Id Firmadoc" val una cosa
# diferent a cada document.
#
# $marcatPerDefecte: si els punts surten ja marcats. Al bloc DESPRES si (el Word
# de l'usuari els portava tots i ell hi anava esborrant el que no tocava; picar
# quinze caselles cada vegada era feina de mes), i al bloc ABANS no, perque alli
# cada punt demana a mes decidir si es te la documentacio o no.
#
# Retorna @{ Nav; Punts; Memoria }.
function Select-LlicDocumentacio($punts, [string]$titol, [string]$subtitol, [bool]$ambEstat,
                                 [bool]$marcatPerDefecte = $false, [bool]$ambDades = $false,
                                 $preSel = $null, [bool]$ambSubs = $false,
                                 [string]$estatPerDefecte = 'no') {
    $punts = @($punts)
    $grups = @(_LlicAgrupaPunts $punts)

    # Estat de cada punt (viu tota la pantalla i es el que es retorna).
    $st = @{}
    for ($i = 0; $i -lt $punts.Count; $i++) {
        $p = $punts[$i]
        $clau = _LlicClauPunt $p
        # Camps  = els objectes de camp VIUS de la pantalla (els fa _RenderRichInto).
        # Valors = el mapa pla nom -> valor, que es el que es RECORDA i es desa a
        #          la base de dades. Els objectes de camp no sobreviuen un pas per
        #          JSON; el mapa pla si, i _RenderRichInto ja el sap llegir com a
        #          $preload (_GetPreloadValue, Camps.ps1).
        $ini = if ([string]::IsNullOrWhiteSpace($estatPerDefecte)) { 'no' } else { [string]$estatPerDefecte }
        # EstatPrevi: el que deia la memoria, tal qual. Una pantalla que NO
        # pregunta l'estat (requeriment, favorable pre) el torna a desar com
        # l'ha trobat: si hi desava el seu 'no' per defecte, el POST sortia amb
        # "No es disposa" marcat a tots els punts en lloc del seu 'si'.
        $e = @{ Marcat = $marcatPerDefecte; Estat = $ini; EstatPrevi = ''; Camps = [ordered]@{}; Valors = @{}; Subs = @{} }
        if ($null -ne $preSel -and $preSel.Contains($clau)) {
            $e.Marcat = [bool]$preSel[$clau].Marcat
            $e.EstatPrevi = [string]$preSel[$clau].Estat
            # Buit = la pantalla d'on ve no el preguntava: mana el de la fase.
            if (-not [string]::IsNullOrWhiteSpace($e.EstatPrevi)) { $e.Estat = $e.EstatPrevi }
            if ($null -ne $preSel[$clau].Valors) { $e.Valors = $preSel[$clau].Valors }
            if ($null -ne $preSel[$clau].Subs)   { $e.Subs   = $preSel[$clau].Subs }
        }
        # Per defecte, TOTS els sub-punts d'un punt marcat entren.
        foreach ($k in 0..([Math]::Max(0, @($p.Subs).Count - 1))) {
            if (-not $e.Subs.Contains($k)) { $e.Subs[$k] = $true }
        }
        $st[$i] = $e
    }

    # LES FUNCIONS DE LA PANTALLA, TOTES DINS D'UN HASHTABLE.
    #
    # PER QUE: .GetNewClosure() copia el VALOR de les variables en el moment de
    # crear el scriptblock. Un scriptblock que es cridi a si mateix (o que
    # cridi un que encara no existeix) es quedaria amb $null i peta amb
    #   "L'expressio que segueix a & ... no es un nom d'ordre ni un scriptblock".
    # El hashtable, en canvi, es captura per REFERENCIA: $fn.Pinta es resol en
    # cridar-lo i l'ordre de definicio deixa d'importar.
    # Hi ha una prova que ho vigila a run-tests.ps1 ("cap closure es refereix a
    # si mateixa"); no tornis a fer $x = { ... & $x ... }.GetNewClosure().
    $fn = @{}
    $estatUi = @{ Busy = $false }

    $form = _NewForm
    $form.Text = $titol
    $form.ClientSize = New-Object System.Drawing.Size(1080, 660)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(820, 520)

    # ---- Esquerra: cercador + ARBRE amb caselles ---------------------------
    $panEsq = New-Object System.Windows.Forms.Panel
    $panEsq.Location = New-Object System.Drawing.Point(14, 66)
    $panEsq.Size = New-Object System.Drawing.Size(500, 520)
    $panEsq.Anchor = 'Top,Bottom,Left'
    [void]$form.Controls.Add($panEsq)

    # Mateix aspecte que el Pas 3 (Select-Items): seccions en negreta i punts a
    # dins. El font BASE es la negreta mes ampla, si no WinForms retalla els
    # nodes que tenen un NodeFont mes ample que el del control.
    $arbre = New-Object System.Windows.Forms.TreeView
    $arbre.Location = New-Object System.Drawing.Point(0, 28)
    $arbre.Size = New-Object System.Drawing.Size(500, 492)
    $arbre.Anchor = 'Top,Bottom,Left,Right'
    $arbre.CheckBoxes = $true
    $arbre.HideSelection = $false
    $arbre.ShowNodeToolTips = $true
    $arbre.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    [void]$panEsq.Controls.Add($arbre)

    # ---- Dreta: detall del punt seleccionat --------------------------------
    $panDret = New-Object System.Windows.Forms.Panel
    $panDret.Location = New-Object System.Drawing.Point(526, 66)
    $panDret.Size = New-Object System.Drawing.Size(540, 520)
    $panDret.Anchor = 'Top,Bottom,Left,Right'
    $panDret.AutoScroll = $true
    $panDret.BorderStyle = 'FixedSingle'
    $panDret.BackColor = [System.Drawing.Color]::White
    [void]$form.Controls.Add($panDret)

    # Reconstrueix l'arbre segons el filtre. L'estat de les caselles NO viu a
    # l'arbre sino a $st: aixi el filtre no en pot perdre cap.
    $fn.Omple = {
        param($filtre)
        $estatUi.Busy = $true
        $arbre.BeginUpdate()
        try {
            $arbre.Nodes.Clear()
            $f = ([string]$filtre).Trim()
            # DOS NIVELLS, com el Pas 3: seccio en negreta i, si en te, la
            # subseccio subratllada a dins. El node de seccio es REAPROFITA
            # entre subseccions consecutives de la mateixa seccio.
            $secAra = [char]0x0001   # cap seccio encara (no pot coincidir amb res)
            $nodeSec = $null
            foreach ($g in $grups) {
                $secTit = [string]$g.Titol
                $subTit = [string]$g.Sub
                $secMatch = ((_TextMatches $secTit $f) -or (_TextMatches $subTit $f))
                # Els punts del grup que passen el filtre.
                $visibles = New-Object System.Collections.ArrayList
                foreach ($i in @($g.Idx)) {
                    $et = _LlicEtiquetaPunt $punts[$i]
                    if ($secMatch -or (_TextMatches $et $f)) { [void]$visibles.Add($i) }
                }
                if ($visibles.Count -eq 0) { continue }

                # Grup sense titol = primer nivell, sense capcalera.
                $pare = $null
                if (-not [string]::IsNullOrWhiteSpace($secTit)) {
                    if ($secTit -ne $secAra) {
                        $secAra = $secTit
                        $nodeSec = New-Object System.Windows.Forms.TreeNode($secTit)
                        $nodeSec.Tag = @{ Kind = 'Section' }
                        $nodeSec.NodeFont = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
                        [void]$arbre.Nodes.Add($nodeSec)
                    }
                    $pare = $nodeSec
                    if (-not [string]::IsNullOrWhiteSpace($subTit)) {
                        $nodeSub = New-Object System.Windows.Forms.TreeNode($subTit)
                        $nodeSub.Tag = @{ Kind = 'Section' }
                        $nodeSub.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Underline)
                        [void]$nodeSec.Nodes.Add($nodeSub)
                        $pare = $nodeSub
                    }
                }
                $totsMarcats = $true
                foreach ($i in $visibles) {
                    $nd = New-Object System.Windows.Forms.TreeNode((_LlicEtiquetaPunt $punts[$i]))
                    $nd.Tag = @{ Kind = 'Item'; Idx = $i }
                    $nd.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
                    $nd.ToolTipText = (_LlicTextPlaDelCos $punts[$i].Cos)
                    if ($nd.ToolTipText.Length -gt 600) { $nd.ToolTipText = $nd.ToolTipText.Substring(0, 600) + '...' }
                    $nd.Checked = [bool]$st[$i].Marcat
                    if (-not $nd.Checked) { $totsMarcats = $false }
                    if ($null -eq $pare) { [void]$arbre.Nodes.Add($nd) } else { [void]$pare.Nodes.Add($nd) }
                }
                if ($null -ne $pare) {
                    $pare.Checked = $totsMarcats
                    $pare.Expand()
                    if ($null -ne $pare.Parent) { $pare.Parent.Expand() }
                }
            }
        } finally {
            $arbre.EndUpdate()
            $estatUi.Busy = $false
        }
    }.GetNewClosure()

    $cerca = _AddSearchBox $panEsq 0 2 380 'Cerca:' {
        param($sender, $ev)
        & $fn.Omple $sender.Text
    }.GetNewClosure()

    $fn.Pinta = {
        param($idx)
        $panDret.Controls.Clear()
        if ($null -eq $idx -or $idx -lt 0) { return }
        $p = $punts[$idx]
        $e = $st[$idx]
        $y = 10
        # UN REGISTRE DE CAMPS NOU A CADA PINTADA, mai un de tota la pantalla.
        #
        # El registre (Camps.ps1) SINCRONITZA els controls que porten el MATEIX
        # nom de camp: escriure en un "Id Firmadoc" copia el text a tots els
        # altres "Id Firmadoc" que hi hagi registrats. Es el que vol REQ1, on un
        # camp val el mateix a tot l'informe. Aqui NO: cada punt es un document
        # diferent. Amb un registre de tota la pantalla, els quadres dels punts
        # ja visitats (trets del panell pero vius, amb el seu handler) rebien
        # el text del punt nou i l'escrivien al SEU punt: l'informe del GIA 924
        # va sortir amb el mateix Id Firmadoc (9887463) als cinc punts d'ABANS.
        # Hi ha un guard que ho vigila (06-guards.ps1).
        $fldRegistry = _NewFieldRegistry

        # El text sencer del punt.
        $lbT = New-Object System.Windows.Forms.Label
        $lbT.Location = New-Object System.Drawing.Point(10, $y)
        $lbT.MaximumSize = New-Object System.Drawing.Size(495, 0)
        $lbT.AutoSize = $true
        $lbT.Text = (_LlicTextPlaDelCos $p.Cos)
        if ([string]::IsNullOrWhiteSpace($lbT.Text)) { $lbT.Text = (_LlicEtiquetaPunt $p 0) }
        [void]$panDret.Controls.Add($lbT)
        $y += [Math]::Max(24, $lbT.PreferredHeight + 10)

        if ($ambEstat) {
            $rbNo = New-Object System.Windows.Forms.RadioButton
            $rbNo.Location = New-Object System.Drawing.Point(10, $y)
            $rbNo.AutoSize = $true
            $rbNo.Text = 'No es disposa del document'
            $rbNo.Checked = ([string]$e.Estat -ne 'si')
            [void]$panDret.Controls.Add($rbNo)
            $y += 24
            $rbSi = New-Object System.Windows.Forms.RadioButton
            $rbSi.Location = New-Object System.Drawing.Point(10, $y)
            $rbSi.AutoSize = $true
            $rbSi.Text = 'Es disposa del document'
            $rbSi.Checked = ([string]$e.Estat -eq 'si')
            [void]$panDret.Controls.Add($rbSi)
            $y += 30
            # UNA COPIA LOCAL DE $fn.
            #
            # .GetNewClosure() nomes copia els LOCALS del context que la crida.
            # Aqui dins, $idx, $e i els dos radios SI que ho son, pero $fn ve del
            # modul de la closure de fora i arribaria als handlers com a $null
            # (-> "& $null.Pinta", el quadre d'error en triar "Es disposa").
            # Hi ha una prova que ho vigila; vegeu CLAUDE.md.
            $fnAquest = $fn
            # UN HANDLER PER RADIO, i nomes actua el que s'acaba de marcar: un
            # sol clic dispara DOS esdeveniments -el que es marca i el germa que
            # es desmarca- i amb un handler compartit la pantalla es repintava
            # dues vegades, la segona llegint uns controls que Controls.Clear()
            # acabava de treure del panell.
            $rbSi.add_CheckedChanged({
                if (-not $rbSi.Checked) { return }
                $e.Estat = 'si'
                & $fnAquest.Pinta $idx
            }.GetNewClosure())
            $rbNo.add_CheckedChanged({
                if (-not $rbNo.Checked) { return }
                $e.Estat = 'no'
                & $fnAquest.Pinta $idx
            }.GetNewClosure())
        }

        # La frase del cataleg amb els camps INLINE (nomes al bloc ABANS).
        if ($ambDades) {
            # TOT EL BLOC JUNT, no linia a linia: un [CAMP:]/[OPCIO:] pot ocupar
            # dos paragrafs del cataleg, i la pantalla ha de veure el mateix
            # text que el generador (que resol per bloc, Apply-FieldsToLines).
            $linies = if ([string]$e.Estat -eq 'si') { @($p.SiDisposa) } else { @($p.NoDisposa) }
            $linies = @(($linies -join [char]10))
            foreach ($l in $linies) {
                if ([string]::IsNullOrWhiteSpace($l)) { continue }
                $flow = New-Object System.Windows.Forms.FlowLayoutPanel
                $flow.Location = New-Object System.Drawing.Point(10, $y)
                $flow.Size = New-Object System.Drawing.Size(500, 10)
                $flow.AutoSize = $true
                $flow.AutoSizeMode = 'GrowAndShrink'
                $flow.MaximumSize = New-Object System.Drawing.Size(500, 0)
                $flow.WrapContents = $true
                $flow.FlowDirection = 'LeftToRight'
                [void]$panDret.Controls.Add($flow)
                # LA MATEIXA funcio que REQ1, amb un diccionari PER PUNT i amb
                # els valors recordats com a $preload: aixi els Id Firmadoc i
                # els expedients de l'informe anterior ja surten escrits.
                _RenderRichInto $flow ([string]$l) $e.Camps $e.Valors $fldRegistry
                $y += [Math]::Max(26, $flow.PreferredSize.Height + 8)
            }
        }

        # Els SUB-PUNTS (nomes al bloc DESPRES): no tothom els te tots.
        if ($ambSubs -and @($p.Subs).Count -gt 0) {
            $lbS = New-Object System.Windows.Forms.Label
            $lbS.Location = New-Object System.Drawing.Point(10, $y)
            $lbS.AutoSize = $true
            $lbS.Text = 'Quins hi entren:'
            [void]$panDret.Controls.Add($lbS)
            $y += 22
            for ($k = 0; $k -lt @($p.Subs).Count; $k++) {
                $sub = @($p.Subs)[$k]
                $txtSub = (@($sub) -join ' ').Trim()
                if ([string]::IsNullOrWhiteSpace($txtSub)) { continue }
                $cb = New-Object System.Windows.Forms.CheckBox
                $cb.Location = New-Object System.Drawing.Point(24, $y)
                $cb.MaximumSize = New-Object System.Drawing.Size(470, 0)
                $cb.AutoSize = $true
                $cb.Text = $txtSub
                $cb.Checked = [bool]$e.Subs[$k]
                $kk = $k
                $cb.add_CheckedChanged({ $e.Subs[$kk] = [bool]$cb.Checked }.GetNewClosure())
                [void]$panDret.Controls.Add($cb)
                $y += [Math]::Max(24, $cb.PreferredHeight + 4)
            }
        }
    }.GetNewClosure()

    # Marcar una SECCIO marca tots els seus punts (com al Pas 3).
    $arbre.add_AfterCheck({
        param($sender, $ev)
        if ($estatUi.Busy) { return }
        $estatUi.Busy = $true
        try {
            $tag = $ev.Node.Tag
            if ($null -ne $tag -and [string]$tag.Kind -eq 'Section') {
                # Baixa per tot l'arbre: una seccio pot tenir subseccions.
                $pila = New-Object System.Collections.ArrayList
                [void]$pila.Add($ev.Node)
                while ($pila.Count -gt 0) {
                    $nd = $pila[0]; [void]$pila.RemoveAt(0)
                    foreach ($fill in $nd.Nodes) {
                        $fill.Checked = $ev.Node.Checked
                        if ($null -ne $fill.Tag -and [string]$fill.Tag.Kind -eq 'Item') {
                            $st[[int]$fill.Tag.Idx].Marcat = [bool]$ev.Node.Checked
                        } else { [void]$pila.Add($fill) }
                    }
                }
            } elseif ($null -ne $tag -and [string]$tag.Kind -eq 'Item') {
                $st[[int]$tag.Idx].Marcat = [bool]$ev.Node.Checked
                # La casella de la seccio segueix els seus fills.
                $pare = $ev.Node.Parent
                while ($null -ne $pare) {
                    $tots = $true
                    foreach ($fill in $pare.Nodes) { if (-not $fill.Checked) { $tots = $false; break } }
                    $pare.Checked = $tots
                    $pare = $pare.Parent
                }
            }
        } finally { $estatUi.Busy = $false }
    }.GetNewClosure())

    $arbre.add_AfterSelect({
        param($sender, $ev)
        if ($estatUi.Busy) { return }
        $tag = $ev.Node.Tag
        if ($null -eq $tag -or [string]$tag.Kind -ne 'Item') { $panDret.Controls.Clear(); return }
        & $fn.Pinta ([int]$tag.Idx)
    }.GetNewClosure())

    & $fn.Omple ''
    if ($arbre.Nodes.Count -gt 0) {
        $primer = $arbre.Nodes[0]
        if ($null -ne $primer.Tag -and [string]$primer.Tag.Kind -ne 'Item' -and $primer.Nodes.Count -gt 0) {
            $primer = $primer.Nodes[0]
        }
        $arbre.SelectedNode = $primer
    }

    # ---- Botons -----------------------------------------------------------
    $res = @{ Nav = 'back'; Punts = @(); Memoria = $null }
    $peu = _AddPeuBotons $form @(
        @{ Nom = 'Enrere'; Text = (_TxtEnrere) },
        @{ Nom = 'Tot'; Text = 'Marcar-ho tot' },
        @{ Nom = 'Cap'; Text = 'Desmarcar-ho tot' }) @(
        @{ Nom = 'Ok'; Text = 'Continuar'; Estil = 'primari' }) 606 -Ancorat
    $btnOk = $peu.Ok; $btnBack = $peu.Enrere; $btnTot = $peu.Tot; $btnCap = $peu.Cap
    $btnOk.add_Click({
        $sel = New-Object System.Collections.ArrayList
        $mem = @{}
        # EN L'ORDRE DEL CATALEG, no el de l'arbre: agrupar es NOMES de pantalla.
        for ($i = 0; $i -lt $punts.Count; $i++) {
            $p = $punts[$i]
            $e = $st[$i]
            $clau = _LlicClauPunt $p
            # El registre de Camps.ps1 desa objectes amb .Value; aqui en volem un
            # mapa nom -> valor, que es el que es recorda i el que es desa.
            # S'HI CONSERVA el que ja hi havia: si un punt no s'ha arribat a
            # pintar (no s'hi ha clicat mai), $e.Camps es buit i els valors
            # recuperats de la base es perdrien.
            $vals = @{}
            foreach ($k in @($e.Valors.Keys)) { $vals[[string]$k] = [string]$e.Valors[$k] }
            foreach ($k in @($e.Camps.Keys))  { $vals[[string]$k] = [string]$e.Camps[$k].Value }
            $e.Valors = $vals
            # Si aquesta pantalla no pregunta l'estat, ni el recorda ni el diu:
            # es desa el que hi havia i el punt surt SENSE estat (cap "No es
            # disposa..." ni "Es disposa..." a l'informe, digui el que digui el
            # cataleg).
            $estatMem = if ($ambEstat) { [string]$e.Estat } else { [string]$e.EstatPrevi }
            $estatPunt = if ($ambEstat) { [string]$e.Estat } else { '' }
            $mem[$clau] = @{ Marcat = $e.Marcat; Estat = $estatMem; Valors = $vals; Subs = $e.Subs }
            if (-not $e.Marcat) { continue }
            $si = @($p.SiDisposa); $no = @($p.NoDisposa)
            if ($estatPunt -eq 'si') { $si = @(_LlicAplicaCamps $p.SiDisposa $vals) }
            else                     { $no = @(_LlicAplicaCamps $p.NoDisposa $vals) }
            # Nomes els sub-punts triats.
            $subs = New-Object System.Collections.ArrayList
            for ($k = 0; $k -lt @($p.Subs).Count; $k++) {
                if ($ambSubs -and -not [bool]$e.Subs[$k]) { continue }
                [void]$subs.Add(@($p.Subs)[$k])
            }
            # 'Select-Object *' i no una llista de camps: el punt porta tambe la
            # SECCIO, la SUBSECCIO i l'INTRO de REQ1, i enumerar-los aqui vol dir
            # que el dia que se n'afegeixi un es perdi en silenci. Ja va passar.
            $c = $p | Select-Object *
            $c.NoDisposa = $no
            $c.SiDisposa = $si
            $c.Subs = $subs.ToArray()
            $c | Add-Member NoteProperty Estat $estatPunt -Force
            [void]$sel.Add($c)
        }
        $res.Punts = $sel.ToArray()
        $res.Memoria = $mem
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())

    $btnBack.add_Click({ $form.Close() }.GetNewClosure())

    $fn.MarcaTot = {
        param($valor)
        for ($i = 0; $i -lt $punts.Count; $i++) { $st[$i].Marcat = $valor }
        & $fn.Omple $cerca.Text
    }.GetNewClosure()
    $btnTot.add_Click({ & $fn.MarcaTot $true }.GetNewClosure())

    $btnCap.add_Click({ & $fn.MarcaTot $false }.GetNewClosure())

    [void](_AddBrandHeader $form $titol $subtitol 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

function Select-LlicTecnic($pre, $preDocs = $null) {
    $form = _NewForm
    $form.Text = 'Llic' + [char]0x00E8 + 'ncia - Documentaci' + [char]0x00F3
    $form.ClientSize = New-Object System.Drawing.Size(620, 430)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $camps = @(
        @{ K = 'Tecnic';  L = 'T' + [char]0x00E8 + 'cnic redactor:' }
        @{ K = 'NumCol';  L = 'N' + [char]0x00FA + 'm. col' + [char]0x00B7 + 'legiat:' }
        @{ K = 'Collegi'; L = 'Col' + [char]0x00B7 + 'legi:' }
        @{ K = 'Data';    L = 'Data de signatura:' }
    )
    $tb = @{}
    $y = 76
    foreach ($c in $camps) {
        $l = New-Object System.Windows.Forms.Label
        $l.Location = New-Object System.Drawing.Point(20, ($y + 3))
        $l.Size = New-Object System.Drawing.Size(150, 20)
        $l.Text = [string]$c.L
        [void]$form.Controls.Add($l)
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location = New-Object System.Drawing.Point(175, $y)
        $t.Size = New-Object System.Drawing.Size(420, 22)
        if ($null -ne $pre -and $pre.Contains([string]$c.K)) { $t.Text = [string]$pre[[string]$c.K] }
        [void]$form.Controls.Add($t)
        $tb[[string]$c.K] = $t
        $y += 32
    }

    $lblD = New-Object System.Windows.Forms.Label
    $lblD.Location = New-Object System.Drawing.Point(20, ($y + 8))
    $lblD.Size = New-Object System.Drawing.Size(560, 20)
    $lblD.Text = 'Quins documents s' + [char]0x2019 + 'han signat, i el seu Id Firmadoc:'
    [void]$form.Controls.Add($lblD)
    $y += 32

    $docs = @(_LlicDocsSignats)
    # EL QUE JA S'HAVIA TRIAT. Sense aixo, tornar Enrere o fer el segon informe
    # de la mateixa llicencia obligava a tornar a marcar-ho i a reescriure els
    # Id Firmadoc.
    $preD = ConvertTo-Mapa $preDocs
    $cbDoc = @{}
    $tbDoc = @{}
    foreach ($d in $docs) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Location = New-Object System.Drawing.Point(30, ($y + 2))
        $cb.Size = New-Object System.Drawing.Size(110, 22)
        $cb.Text = [string]$d
        [void]$form.Controls.Add($cb)
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location = New-Object System.Drawing.Point(150, $y)
        $t.Size = New-Object System.Drawing.Size(300, 22)
        [void]$form.Controls.Add($t)
        $lid = New-Object System.Windows.Forms.Label
        $lid.Location = New-Object System.Drawing.Point(458, ($y + 3))
        $lid.Size = New-Object System.Drawing.Size(140, 20)
        $lid.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
        $lid.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $lid.Text = 'Id Firmadoc'
        [void]$form.Controls.Add($lid)
        if ($preD.ContainsKey([string]$d)) {
            $e = $preD[[string]$d]
            if ($e -is [System.Collections.IDictionary]) { $cb.Checked = [bool]$e['Marcat']; $t.Text = [string]$e['Id'] }
            elseif ($null -ne $e) { $cb.Checked = [bool]$e.Marcat; $t.Text = [string]$e.Id }
        }
        $cbDoc[[string]$d] = $cb
        $tbDoc[[string]$d] = $t
        $y += 30
    }

    $res = @{ Nav = 'back'; Text = ''; Items = @(); Camps = @{}; Docs = (_LlicDocsBuits) }
    $peu = _AddPeuBotons $form @(@{ Nom = 'Enrere'; Text = (_TxtEnrere) }) @(
        @{ Nom = 'Ok'; Text = 'Continuar'; Estil = 'primari' }) 380
    $btnOk = $peu.Ok; $btnBack = $peu.Enrere
    $btnOk.add_Click({
        $res.Text = _LlicTextDocumentacio $tb['Tecnic'].Text $tb['NumCol'].Text $tb['Collegi'].Text $tb['Data'].Text
        $tria = [ordered]@{}
        foreach ($d in $docs) {
            $tria[[string]$d] = @{
                Marcat = [bool]$cbDoc[[string]$d].Checked
                Id     = ([string]$tbDoc[[string]$d].Text).Trim()
            }
        }
        $res.Docs = $tria
        $res.Items = @(_LlicItemsDocsSignats $tria)
        foreach ($k in $tb.Keys) { $res.Camps[$k] = [string]$tb[$k].Text }
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())

    $btnBack.add_Click({ $form.Close() }.GetNewClosure())

    [void](_AddBrandHeader $form ('Documentaci' + [char]0x00F3) ('Qui ha signat el projecte i amb quin Id Firmadoc') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

# Pas de les CONDICIONS (nomes als favorables): QUINS ACTORS les posen i, de
# cada un, el PDF del seu informe.
#
# Abans era un quadre de text lliure (i despres, una casella al pas 1). L'usuari
# va explicar que les condicions les posen els organismes que informen els
# punts d'Autoritzacions / Informes preceptius, i que el que ha de dir l'informe
# es QUINS: els seus informes van adjunts darrere. Una llista per marcar, i amb
# un de marcat ja hi ha condicions.
#
# $pdfs: nom -> ruta ja triada (la copia local, si ve de la memoria).
# Retorna @{ Nav; Actors (els noms marcats, en l'ordre de la llista); Pdfs }.
function Select-LlicCondicions($actors, $marcats, $pdfs = $null) {
    $actors = @($actors)
    $form = _NewForm
    $form.Text = 'Condicions de la llic' + [char]0x00E8 + 'ncia'
    $form.ClientSize = New-Object System.Drawing.Size(700, 500)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 70)
    $lbl.Size = New-Object System.Drawing.Size(660, 52)
    $lbl.Text = ('Marca qui posa condicions i tria el PDF del seu informe: en passar l' + [char]0x2019 + 'informe a PDF ' +
                 's' + [char]0x2019 + 'hi afegira darrere. Si no en marques cap, la conclusi' + [char]0x00F3 +
                 ' no parla de condicions. Surten marcats els que tenen l' + [char]0x2019 + 'informe preceptiu com a "Es disposa".')
    [void]$form.Controls.Add($lbl)

    $pan = New-Object System.Windows.Forms.Panel
    $pan.Location = New-Object System.Drawing.Point(20, 128)
    $pan.Size = New-Object System.Drawing.Size(660, 302)
    $pan.Anchor = 'Top,Bottom,Left,Right'
    $pan.AutoScroll = $true
    $pan.BorderStyle = 'FixedSingle'
    $pan.BackColor = [System.Drawing.Color]::White
    [void]$form.Controls.Add($pan)

    $marcatsSet = @{}
    foreach ($m in @($marcats)) { $marcatsSet[([string]$m).Trim().ToLowerInvariant()] = $true }
    $files = New-Object System.Collections.ArrayList
    $y = 8
    foreach ($a in $actors) {
        $nom = [string]$a.Nom
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Location = New-Object System.Drawing.Point(8, ($y + 2))
        $cb.Size = New-Object System.Drawing.Size(290, 22)
        $cb.AutoEllipsis = $true
        $cb.Text = $nom
        $cb.Checked = $marcatsSet.ContainsKey($nom.Trim().ToLowerInvariant())
        [void]$pan.Controls.Add($cb)
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(304, $y)
        $tb.Size = New-Object System.Drawing.Size(250, 24)
        if ($null -ne $pdfs -and $pdfs.Contains($nom)) { $tb.Text = [string]$pdfs[$nom] }
        [void]$pan.Controls.Add($tb)
        $bt = New-Object System.Windows.Forms.Button
        $bt.Location = New-Object System.Drawing.Point(560, ($y - 1))
        $bt.Size = New-Object System.Drawing.Size(70, 26)
        $bt.Text = 'PDF' + [char]0x2026
        _StyleSecondaryButton $bt
        # Triar un PDF marca l'actor: si t'hi has molestat, es que hi va.
        $bt.add_Click({
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Filter = 'PDF (*.pdf)|*.pdf'
            $dlg.Title = 'Informe de: ' + $cb.Text
            try {
                $dir = Split-Path -Parent $tb.Text
                if ($dir -and (Test-Path -LiteralPath $dir)) { $dlg.InitialDirectory = $dir }
            } catch { }
            if ($dlg.ShowDialog() -eq 'OK') { $tb.Text = $dlg.FileName; $cb.Checked = $true }
            $dlg.Dispose()
        }.GetNewClosure())
        [void]$pan.Controls.Add($bt)
        [void]$files.Add(@{ Nom = $nom; Cb = $cb; Tb = $tb })
        $y += 32
    }

    $res = @{ Nav = 'back'; Actors = @(); Pdfs = @{} }
    $peu = _AddPeuBotons $form @(@{ Nom = 'Enrere'; Text = (_TxtEnrere) }) @(
        @{ Nom = 'Ok'; Text = 'Continuar'; Estil = 'primari' }) 448 -Ancorat
    $btnOk = $peu.Ok; $btnBack = $peu.Enrere
    $btnOk.add_Click({
        # En l'ordre de la LLISTA (el del cataleg), no en el que s'han clicat.
        $tri = New-Object System.Collections.ArrayList
        $rutes = @{}
        $sensePdf = New-Object System.Collections.ArrayList
        foreach ($f in $files) {
            $ruta = ([string]$f.Tb.Text).Trim().Trim('"')
            # Es recorden TOTES les rutes, tambe les dels no marcats: desmarcar
            # un actor un moment no ha de fer perdre el seu PDF.
            if ($ruta) { $rutes[[string]$f.Nom] = $ruta }
            if (-not $f.Cb.Checked) { continue }
            [void]$tri.Add([string]$f.Nom)
            if (-not $ruta) { [void]$sensePdf.Add([string]$f.Nom) }
        }
        if ($sensePdf.Count -gt 0) {
            $r = [System.Windows.Forms.MessageBox]::Show(
                ("D'aquests no has triat el PDF:`n`n  " + ($sensePdf -join "`n  ") + "`n`n" +
                 "L'informe els anomenara, pero quan el passis a PDF no s'hi adjuntara el seu informe.`n`nContinuar igualment?"),
                'Condicions', 'YesNo', 'Warning')
            if ($r -ne 'Yes') { return }
        }
        $res.Actors = $tri.ToArray()
        $res.Pdfs = $rutes
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())

    $btnBack.add_Click({ $form.Close() }.GetNewClosure())

    [void](_AddBrandHeader $form 'Condicions' ('Qui les posa i el seu informe (nom' + [char]0x00E9 + 's als favorables)') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}
