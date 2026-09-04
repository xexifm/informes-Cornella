#requires -Version 5.1
<#
.SYNOPSIS
  Les dues pantalles de TRIA de l'assistent: les deficiencies i les conclusions.

.DESCRIPTION
  - Select-Items (Pas 3): arbre amb caselles i filtre de text sobre el cataleg.
    L'estructura ($sections) es manté a part i l'arbre es reconstrueix quan
    canvia el filtre, conservant el que ja tenies marcat.
  - Select-Conclusions (Pas 5): tria de conclusions, amb els camps
    [CAMP:]/[OPCIO:] incrustats dins del mateix text.

  Ve de Motor.ps1. Es dot-sourceja des d'alli: mateix ambit, mateix
  comportament.
#>

# ----------------------------------------------------------------------------
# Step 3 (UI) - TreeView amb filtre + checkboxes
# ----------------------------------------------------------------------------
# Helpers per al filtre del TreeView. Es manté l'estructura $sections a part i
# es reconstrueix el tree quan canvia el filtre, preservant els check states.

function _TextMatches($text, $needle) {
    if ([string]::IsNullOrEmpty($needle)) { return $true }
    if ($null -eq $text) { return $false }
    return $text.ToLower().Contains($needle.ToLower())
}

# Reconstrueix el TreeView segons el text de filtre. Preserva check states
# (passats en una hashtable [key] -> bool) i els actualitza durant la construccio.
# Bloquegem la propagacio automatica de check durant el rebuild perque
# marcar nodes programmaticament dispara l'event AfterCheck.
# Construeix el text del tooltip d'un item (o fill) del TreeView del Pas 3:
# concatena les BodyLines descartant els enllacos (URL-only i URLs incrustats)
# perque l'usuari pugui veure el text de l'item en passar el ratolí. Aplica
# una mica de neteja per fer-lo llegible i el limita a 600 caracters per
# evitar tooltips desmesurats.
function _GetItemTooltip($el) {
    if ($null -eq $el -or $null -eq $el.BodyLines) { return '' }
    $parts = New-Object System.Collections.ArrayList
    foreach ($ln in $el.BodyLines) {
        $s = [string]$ln
        if ([string]::IsNullOrWhiteSpace($s)) { continue }
        # Linies marcades per Cita: tota la linia es URL.
        if ($s.StartsWith('[[URL]] ')) { continue }
        # Linies que son nomes URL.
        if ($s.Trim() -match '^https?://') { continue }
        # Linia mixta: extreu nomes el text (descarta URLs incrustats).
        $p = _SplitTextAndUrls $s
        if (-not [string]::IsNullOrWhiteSpace($p.Text)) { [void]$parts.Add($p.Text) }
    }
    $tip = ($parts -join [Environment]::NewLine)
    if ($tip.Length -gt 600) { $tip = $tip.Substring(0, 600) + '...' }
    return $tip
}

function _RebuildTree($tv, $sections, $needle, $checkStates) {
    $tv.BeginUpdate()
    $script:_propagating = $true
    try {
        $tv.Nodes.Clear()
        foreach ($sec in $sections) {
            $secMatches = _TextMatches $sec.Title $needle

            # Recollim items/children que cal mostrar
            $itemNodesToAdd = New-Object System.Collections.ArrayList
            $currentContainer = $null
            foreach ($el in $sec.Items) {
                if ($el.Kind -eq 'subsection') {
                    $subShow = $secMatches -or (_TextMatches $el.Short $needle)
                    $itemNodesToAdd.Add(@{ Kind='Subsection'; El=$el; ChildShows=@(); ShowMe=$subShow }) | Out-Null
                    continue
                }
                if ($el.Kind -eq 'intro') { continue }  # mai al TreeView

                $itemMatches = _TextMatches $el.Short $needle
                $matchedChildren = New-Object System.Collections.ArrayList
                foreach ($ch in $el.Children) {
                    if ($secMatches -or $itemMatches -or (_TextMatches $ch.Short $needle)) {
                        [void]$matchedChildren.Add($ch)
                    }
                }
                $showItem = $secMatches -or $itemMatches -or ($matchedChildren.Count -gt 0)
                if ($showItem) {
                    $itemNodesToAdd.Add(@{ Kind='Item'; El=$el; ChildShows=$matchedChildren; ShowMe=$true }) | Out-Null
                }
            }

            # Si el filtre no es buit i no hi ha cap item/subsection visible,
            # ometem la seccio del tot (tret que el titol de la seccio matchi).
            $anyChild = $false
            foreach ($n in $itemNodesToAdd) { if ($n.ShowMe) { $anyChild = $true; break } }
            if (-not $secMatches -and -not $anyChild) { continue }

            $secNode = New-Object System.Windows.Forms.TreeNode($sec.Title)
            $secNode.Tag = @{ Kind = 'Section'; Ref = $sec; Key = $sec.Title }
            $secNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
            $stKey = "SECT::$($sec.Title)"
            if ($checkStates.ContainsKey($stKey)) { $secNode.Checked = $checkStates[$stKey] }
            [void]$tv.Nodes.Add($secNode)

            $container = $secNode
            foreach ($n in $itemNodesToAdd) {
                if (-not $n.ShowMe) { continue }
                if ($n.Kind -eq 'Subsection') {
                    $subNode = New-Object System.Windows.Forms.TreeNode($n.El.Short)
                    $subNode.Tag = @{ Kind = 'Subsection'; Ref = $n.El }
                    $subNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Underline)
                    [void]$secNode.Nodes.Add($subNode)
                    $container = $subNode
                    continue
                }
                # Item. NodeFont explicit (regular) perque el font base del
                # TreeView es negreta (vegeu nota a Select-Items) i no volem
                # que els items surtin en negreta ni que es retallin.
                $itNode = New-Object System.Windows.Forms.TreeNode($n.El.Short)
                $itNode.Tag = @{ Kind = 'Item'; Ref = $n.El; SectionTitle = $sec.Title }
                $itNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
                $itNode.ToolTipText = _GetItemTooltip $n.El
                $itKey = (_ItemKey $sec.Title $n.El.Short)
                if ($checkStates.ContainsKey($itKey)) { $itNode.Checked = $checkStates[$itKey] }
                [void]$container.Nodes.Add($itNode)
                foreach ($ch in $n.ChildShows) {
                    $chNode = New-Object System.Windows.Forms.TreeNode($ch.Short)
                    $chNode.Tag = @{ Kind = 'Child'; Ref = $ch; SectionTitle = $sec.Title; ParentShort = $n.El.Short }
                    $chNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
                    $chNode.ToolTipText = _GetItemTooltip $ch
                    $chKey = (_ItemKey $sec.Title $n.El.Short $ch.Short)
                    if ($checkStates.ContainsKey($chKey)) { $chNode.Checked = $checkStates[$chKey] }
                    [void]$itNode.Nodes.Add($chNode)
                }
            }
            $secNode.ExpandAll()
        }
    } finally {
        $script:_propagating = $false
        $tv.EndUpdate()
    }
}

# Recorre el TreeView i llegeix tots els check states en una hashtable.
function _CollectCheckStates($tv, $checkStates) {
    foreach ($secNode in $tv.Nodes) {
        $secTitle = $secNode.Tag.Ref.Title
        $checkStates["SECT::$secTitle"] = [bool]$secNode.Checked
        foreach ($node in $secNode.Nodes) {
            $kind = $node.Tag.Kind
            if ($kind -eq 'Subsection') {
                foreach ($itNode in $node.Nodes) {
                    $itShort = $itNode.Tag.Ref.Short
                    $checkStates[(_ItemKey $secTitle $itShort)] = [bool]$itNode.Checked
                    foreach ($chNode in $itNode.Nodes) {
                        $checkStates[(_ItemKey $secTitle $itShort $chNode.Tag.Ref.Short)] = [bool]$chNode.Checked
                    }
                }
            } elseif ($kind -eq 'Item') {
                $itShort = $node.Tag.Ref.Short
                $checkStates[(_ItemKey $secTitle $itShort)] = [bool]$node.Checked
                foreach ($chNode in $node.Nodes) {
                    $checkStates[(_ItemKey $secTitle $itShort $chNode.Tag.Ref.Short)] = [bool]$chNode.Checked
                }
            }
        }
    }
}

function Select-Items {
    # $permetreBuit: a Llicencia el bloc 'Projecte' pot no tenir cap punt (hi ha
    # activitats sense cap deficiencia de projecte) i no s'ha de bloquejar el pas.
    param($sections, $preloadSelectedKeys = $null, $fields = $null, $preloadValues = $null,
          [bool]$permetreBuit = $false)
    if ($null -eq $fields) { $fields = [ordered]@{} }

    $form = _NewForm
    $form.Text = 'Pas 3 - Seleccio de deficiencies'
    $form.Size = New-Object System.Drawing.Size(1180, 818)
    $form.StartPosition = 'CenterScreen'

    # Espai reservat a dalt per la banda de marca + barra de passos.
    $topOffset = 78

    # Filtre (textbox al capdamunt). Cada vegada que canvia, es reconstrueix
    # el TreeView amb nomes les coincidencies. Els check states es preserven.
    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = 'Filtre:'
    $lblFilter.Location = New-Object System.Drawing.Point(10, (14 + $topOffset))
    $lblFilter.AutoSize = $true
    $form.Controls.Add($lblFilter)

    $tbFilter = New-Object System.Windows.Forms.TextBox
    $tbFilter.Location = New-Object System.Drawing.Point(60, (10 + $topOffset))
    $tbFilter.Size = New-Object System.Drawing.Size(400, 22)
    $tbFilter.Anchor = 'Top, Left'
    $form.Controls.Add($tbFilter)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = 'Esborra'
    $btnClear.Location = New-Object System.Drawing.Point(465, (9 + $topOffset))
    $btnClear.Size = New-Object System.Drawing.Size(70, 24)
    $btnClear.add_Click({ $tbFilter.Text = '' })
    $form.Controls.Add($btnClear)

    $tv = New-Object System.Windows.Forms.TreeView
    $tv.Location = New-Object System.Drawing.Point(10, (40 + $topOffset))
    $tv.Size = New-Object System.Drawing.Size(560, 610)
    $tv.CheckBoxes = $true
    $tv.HideSelection = $false
    $tv.ShowNodeToolTips = $true
    $tv.Anchor = 'Top, Bottom, Left'
    # El font base es la negreta mes ampla que faran servir les seccions. Aixo
    # evita el bug de WinForms en que un node amb NodeFont mes ample que el
    # font del control surt retallat. Items i fills posen NodeFont regular.
    $tv.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($tv)

    # Panell de detall (dreta): mostra el TEXT de les deficiencies marcades amb
    # els seus desplegables/camps inline per omplir-los aqui mateix.
    $lblDetail = New-Object System.Windows.Forms.Label
    $lblDetail.Text = 'Text i opcions de les deficiencies marcades:'
    $lblDetail.Location = New-Object System.Drawing.Point(585, (14 + $topOffset))
    $lblDetail.AutoSize = $true
    $lblDetail.Anchor = 'Top, Left, Right'
    $form.Controls.Add($lblDetail)

    $detailHost = New-Object System.Windows.Forms.Panel
    $detailHost.Location = New-Object System.Drawing.Point(585, (40 + $topOffset))
    $detailHost.Size = New-Object System.Drawing.Size(575, 610)
    $detailHost.AutoScroll = $true
    $detailHost.BorderStyle = 'FixedSingle'
    $detailHost.Anchor = 'Top, Bottom, Left, Right'
    $form.Controls.Add($detailHost)

    $detailV = New-Object System.Windows.Forms.FlowLayoutPanel
    $detailV.FlowDirection = 'TopDown'
    $detailV.WrapContents = $false
    $detailV.AutoSize = $true
    $detailV.AutoSizeMode = 'GrowAndShrink'
    $detailV.Location = New-Object System.Drawing.Point(0, 0)
    $detailV.Width = 550
    $detailHost.Controls.Add($detailV)

    # Estats de check persistents entre rebuilds. Inicialitzat des de session.
    $checkStates = @{}
    if ($preloadSelectedKeys) {
        foreach ($k in $preloadSelectedKeys) {
            if ([string]::IsNullOrWhiteSpace($k)) { continue }
            $checkStates[[string]$k] = $true
        }
    }

    # Amplada util del panell de detall (sense la barra de desplacament).
    $detailInnerWidth = { [Math]::Max(280, $detailHost.ClientSize.Width - 26) }

    # Crea un FlowLayoutPanel que EMBOLCALLA el text (wrap) a l'ample disponible.
    # Clau: amb AutoSize + WrapContents cal limitar l'amplada amb MaximumSize,
    # si no, el panell creix cap a la dreta i no salta de linia.
    $newWrapFlow = {
        param($leftMargin)
        $f = New-Object System.Windows.Forms.FlowLayoutPanel
        $f.FlowDirection = 'LeftToRight'; $f.WrapContents = $true
        $f.AutoSize = $true; $f.AutoSizeMode = 'GrowAndShrink'
        $f.Margin = New-Object System.Windows.Forms.Padding($leftMargin, 0, 2, 4)
        $iw = & $detailInnerWidth
        $f.MaximumSize = New-Object System.Drawing.Size([Math]::Max(120, ($iw - $leftMargin - 6)), 0)
        return $f
    }

    # Reajusta l'amplada maxima de tots els blocs quan la finestra canvia de mida
    # (sense reconstruir res, per no perdre el focus mentre s'escriu).
    $applyDetailWidths = {
        $iw = & $detailInnerWidth
        $detailV.MaximumSize = New-Object System.Drawing.Size($iw, 0)
        foreach ($child in $detailV.Controls) {
            if ($child -is [System.Windows.Forms.FlowLayoutPanel]) {
                $child.MaximumSize = New-Object System.Drawing.Size([Math]::Max(120, ($iw - $child.Margin.Left - 6)), 0)
            }
        }
    }

    # Reconstrueix el panell de detall a partir de les caselles marcades.
    $refreshDetail = {
        _CollectCheckStates $tv $checkStates
        $detailV.SuspendLayout()
        $detailV.Controls.Clear()
        $registry = _NewFieldRegistry
        $detailV.MaximumSize = New-Object System.Drawing.Size((& $detailInnerWidth), 0)
        $anyShown = $false
        foreach ($sec in $sections) {
            foreach ($el in $sec.Items) {
                if ($el.Kind -ne 'item') { continue }
                $itKey = (_ItemKey $sec.Title $el.Short)
                $itSel = $checkStates.ContainsKey($itKey) -and $checkStates[$itKey]
                $selChildren = New-Object System.Collections.ArrayList
                foreach ($ch in $el.Children) {
                    $chKey = (_ItemKey $sec.Title $el.Short $ch.Short)
                    if ($checkStates.ContainsKey($chKey) -and $checkStates[$chKey]) { [void]$selChildren.Add($ch) }
                }
                if (-not $itSel -and $selChildren.Count -eq 0) { continue }
                $anyShown = $true

                $hdr = New-Object System.Windows.Forms.Label
                $hdr.AutoSize = $true
                $hdr.Text = $el.Short
                $hdr.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
                $hdr.Margin = New-Object System.Windows.Forms.Padding(2, 8, 2, 2)
                [void]$detailV.Controls.Add($hdr)

                if ($itSel) {
                    $txt = _RichTextOfBodyLines $el.BodyLines
                    if ($txt) {
                        $flow = & $newWrapFlow 8
                        _RenderRichInto $flow $txt $fields $preloadValues $registry
                        [void]$detailV.Controls.Add($flow)
                    }
                }
                foreach ($ch in $selChildren) {
                    $ctxt = _RichTextOfBodyLines $ch.BodyLines
                    if (-not $ctxt) { continue }
                    $cflow = & $newWrapFlow 24
                    _RenderRichInto $cflow $ctxt $fields $preloadValues $registry
                    [void]$detailV.Controls.Add($cflow)
                }
            }
        }
        if (-not $anyShown) {
            $empty = New-Object System.Windows.Forms.Label
            $empty.AutoSize = $true
            $empty.ForeColor = [System.Drawing.Color]::DimGray
            $empty.Text = 'Marca deficiencies a l''esquerra per veure-les i omplir-ne les opcions.'
            $empty.Margin = New-Object System.Windows.Forms.Padding(6, 10, 2, 2)
            [void]$detailV.Controls.Add($empty)
        }
        $detailV.ResumeLayout()
    }

    # Build inicial sense filtre
    _RebuildTree $tv $sections '' $checkStates

    # Propagacio recursiva: marcar un node marca tots els descendents.
    $script:_propagating = $false
    $propagate = {
        param($node)
        foreach ($c in $node.Nodes) {
            $c.Checked = $node.Checked
            & $propagate $c
        }
    }
    $tv.add_AfterCheck({
        param($sender, $e)
        if ($script:_propagating) { return }
        $script:_propagating = $true
        try { & $propagate $e.Node } finally { $script:_propagating = $false }
        & $refreshDetail
    })

    # Refilter en temps real (debouncing simple: rebuild a cada keystroke;
    # amb 131 items va fluid)
    $tbFilter.add_TextChanged({
        # Guardem l'estat actual ABANS de reconstruir
        _CollectCheckStates $tv $checkStates
        _RebuildTree $tv $sections $tbFilter.Text $checkStates
    })

    # Quan la finestra canvia de mida, reajustem l'amplada dels blocs perque el
    # text es reembolcalli a l'ample nou (sense reconstruir, per no perdre focus).
    $detailHost.add_SizeChanged({ & $applyDetailWidths })

    # Detall inicial (mostra els ja marcats per precarrega).
    & $refreshDetail

    $back = New-Object System.Windows.Forms.Button
    $back.Text = ([char]0x2190 + ' Enrere')
    $back.Location = New-Object System.Drawing.Point(10, 740)
    $back.Size = New-Object System.Drawing.Size(100, 30)
    $back.DialogResult = 'Retry'
    $back.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $back
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = ('Seg' + [char]0x00FC + 'ent ' + [char]0x2192)
    $ok.Location = New-Object System.Drawing.Point(1050, 740)
    $ok.Size = New-Object System.Drawing.Size(110, 30)
    $ok.DialogResult = 'OK'
    $ok.Anchor = 'Bottom, Right'
    _StylePrimaryButton $ok
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    [void](_AddStepBar $form 3)
    [void](_AddBrandHeader $form ("Defici" + [char]0x00E8 + "ncies de l'activitat") $null 44)

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }

    # Recollim l'estat final i el barregem amb el que tenim memoritzat per
    # items que ara mateix no es mostren (perque hi hagi filtre actiu).
    _CollectCheckStates $tv $checkStates

    # Construim el resultat en ordre del data, preservant subseccions/intros.
    $result = _SeccionsTriades $sections $checkStates
    if ($result.Count -eq 0 -and -not $permetreBuit) {
        [System.Windows.Forms.MessageBox]::Show('No s''ha seleccionat cap deficiencia.','Avis','OK','Warning') | Out-Null
        return [pscustomobject]@{ Nav='stay' }   # es torna a mostrar el Pas 3
    }
    return [pscustomobject]@{ Nav='next'; Data=$result }
}

# De l'estat de les caselles a la llista de seccions triades.
#
# PURA, i per aixo es pot provar sense finestres. Estava COPIADA linia a linia
# (42) a Paquet.ps1: la pantalla del Pas 3 i el cami "des d'un paquet del mobil"
# construeixen el mateix model i cap dels dos no ho ha de decidir pel seu compte
# -si divergissin, el mobil i el PC generarien informes diferents amb la mateixa
# tria-. El que si que es de cada crider es QUE fa amb el resultat: la pantalla
# avisa si no s'ha triat res i torna un {Nav; Data}; el paquet el torna i prou.
#
# $checkStates: clau "Seccio::Item[::Fill]" -> $true si esta marcat.
function _SeccionsTriades($sections, $checkStates) {
$result = New-Object System.Collections.ArrayList
foreach ($sec in $sections) {
    $chosen = New-Object System.Collections.ArrayList
    foreach ($el in $sec.Items) {
        if ($el.Kind -in 'subsection','intro') {
            [void]$chosen.Add([pscustomobject]@{
                Kind      = $el.Kind
                Short     = $el.Short
                BodyLines = $el.BodyLines
                Children  = @()
                Selected  = $false
            })
            continue
        }
        $itKey = (_ItemKey $sec.Title $el.Short)
        $isSel = $checkStates.ContainsKey($itKey) -and $checkStates[$itKey]
        $chosenChildren = New-Object System.Collections.ArrayList
        foreach ($ch in $el.Children) {
            $chKey = (_ItemKey $sec.Title $el.Short $ch.Short)
            if ($checkStates.ContainsKey($chKey) -and $checkStates[$chKey]) {
                [void]$chosenChildren.Add($ch)
            }
        }
        if ($isSel -or $chosenChildren.Count -gt 0) {
            [void]$chosen.Add([pscustomobject]@{
                Kind      = 'item'
                Short     = $el.Short
                BodyLines = $el.BodyLines
                Children  = $chosenChildren
                Selected  = [bool]$isSel
            })
        }
    }
    $hasRealItem = $false
    foreach ($x in $chosen) { if ($x.Kind -eq 'item') { $hasRealItem = $true; break } }
    if ($hasRealItem) {
        [void]$result.Add([pscustomobject]@{
            Title = $sec.Title
            Items = $chosen
        })
    }
}
    return $result
}

# Extreu les claus "Seccio::Item[::Fill]" del resultat de Select-Items, per
# desar-les a la sessio.
function Get-SelectedKeysFromResult($selectedSections) {
    $keys = New-Object System.Collections.ArrayList
    foreach ($sec in $selectedSections) {
        foreach ($it in $sec.Items) {
            if ($it.Kind -ne 'item') { continue }
            if ($it.Selected) { [void]$keys.Add((_ItemKey $sec.Title $it.Short)) }
            foreach ($ch in $it.Children) {
                [void]$keys.Add((_ItemKey $sec.Title $it.Short $ch.Short))
            }
        }
    }
    return $keys.ToArray()
}


function Select-Conclusions {
    # $conclusions   : array d'objectes {Title, Body} (de Read-Conclusions).
    # $always        : array de cadenes ::SEMPRE:: (es mostren com a fixes).
    # $fields        : diccionari de camps compartit (s'hi afegeixen/editen els
    #                  [OPCIO:]/[CAMP:] de les conclusions, inline).
    # $preloadTitles : array de titols preseleccionats (sessio anterior).
    # $preloadValues : valors de camps precarregats (sessio anterior).
    param($conclusions, $always = @(), $fields = $null, $preloadTitles = $null, $preloadValues = $null)
    if ($null -eq $fields) { $fields = [ordered]@{} }
    if ($conclusions.Count -eq 0) { return [pscustomobject]@{ Nav='next'; Data=@() } }

    # Convertim preloadTitles a un HashSet per a comparacio rapida.
    $preloadSet = New-Object System.Collections.Generic.HashSet[string]
    if ($preloadTitles) { foreach ($t in $preloadTitles) { [void]$preloadSet.Add([string]$t) } }

    $form = _NewForm
    $form.Text = 'Pas 4 - Conclusions'
    $form.Size = New-Object System.Drawing.Size(940, 738)
    $form.StartPosition = 'CenterScreen'

    # Espai reservat a dalt per la banda de marca + barra de passos.
    $topOffset = 78

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Marca les conclusions a incloure i omple-hi les opcions/camps:'
    $lbl.Location = New-Object System.Drawing.Point(15, (10 + $topOffset))
    $lbl.AutoSize = $true
    $lbl.Anchor = 'Top, Left, Right'
    $form.Controls.Add($lbl)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(15, (35 + $topOffset))
    $panel.Size = New-Object System.Drawing.Size(895, 545)
    $panel.AutoScroll = $true
    $panel.BorderStyle = 'FixedSingle'
    $panel.Anchor = 'Top, Bottom, Left, Right'
    $form.Controls.Add($panel)

    $listV = New-Object System.Windows.Forms.FlowLayoutPanel
    $listV.FlowDirection = 'TopDown'
    $listV.WrapContents = $false
    $listV.AutoSize = $true
    $listV.AutoSizeMode = 'GrowAndShrink'
    $listV.Location = New-Object System.Drawing.Point(0, 0)
    $panel.Controls.Add($listV)

    $registry = _NewFieldRegistry

    # Amplada util (sense barra de desplacament). El cos de cada conclusio
    # s'embolcalla (wrap) a aquesta amplada; cal MaximumSize perque un
    # FlowLayoutPanel amb AutoSize+WrapContents salti de linia en lloc de
    # creixer cap a la dreta.
    $innerW = { [Math]::Max(280, $panel.ClientSize.Width - 26) }
    $mkConclFlow = {
        param($leftMargin)
        $f = New-Object System.Windows.Forms.FlowLayoutPanel
        $f.FlowDirection = 'LeftToRight'; $f.WrapContents = $true
        $f.AutoSize = $true; $f.AutoSizeMode = 'GrowAndShrink'
        $f.Margin = New-Object System.Windows.Forms.Padding($leftMargin, 0, 2, 6)
        $f.MaximumSize = New-Object System.Drawing.Size([Math]::Max(120, ((& $innerW) - $leftMargin - 6)), 0)
        return $f
    }

    $checks = @()
    for ($i = 0; $i -lt $conclusions.Count; $i++) {
        $c = $conclusions[$i]
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $c.Title
        $cb.AutoSize = $true
        $cb.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $cb.Margin = New-Object System.Windows.Forms.Padding(4, 12, 4, 2)
        $cb.Tag = $c
        if ($preloadSet.Contains([string]$c.Title)) { $cb.Checked = $true }
        [void]$listV.Controls.Add($cb)
        $checks += $cb

        $flow = & $mkConclFlow 22
        _RenderRichInto $flow ([string]$c.Body) $fields $preloadValues $registry
        [void]$listV.Controls.Add($flow)
    }

    # Frases fixes (::SEMPRE::): es mostren perque l'usuari les vegi i, si tenen
    # algun [CAMP:]/[OPCIO:], les pugui omplir (al document hi van sempre).
    $alwaysArr = @($always)
    if ($alwaysArr.Count -gt 0) {
        $sep = New-Object System.Windows.Forms.Label
        $sep.Text = 'Es posa sempre al final:'
        $sep.AutoSize = $true
        $sep.ForeColor = [System.Drawing.Color]::DimGray
        $sep.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Italic)
        $sep.Margin = New-Object System.Windows.Forms.Padding(4, 16, 4, 2)
        [void]$listV.Controls.Add($sep)
        foreach ($a in $alwaysArr) {
            $aflow = & $mkConclFlow 22
            _RenderRichInto $aflow ([string]$a) $fields $preloadValues $registry
            [void]$listV.Controls.Add($aflow)
        }
    }

    # En canviar la mida de la finestra, reembolcallem els cossos a l'ample nou.
    $panel.add_SizeChanged({
        $iw = & $innerW
        $listV.MaximumSize = New-Object System.Drawing.Size($iw, 0)
        foreach ($child in $listV.Controls) {
            if ($child -is [System.Windows.Forms.FlowLayoutPanel]) {
                $child.MaximumSize = New-Object System.Drawing.Size([Math]::Max(120, ($iw - $child.Margin.Left - 6)), 0)
            }
        }
    })

    $back = New-Object System.Windows.Forms.Button
    $back.Text = ([char]0x2190 + ' Enrere')
    $back.Location = New-Object System.Drawing.Point(15, 670)
    $back.Size = New-Object System.Drawing.Size(100, 30)
    $back.DialogResult = 'Retry'
    $back.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $back
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = ('Generar informe ' + [char]0x2192)
    $ok.Location = New-Object System.Drawing.Point(775, 670)
    $ok.Size = New-Object System.Drawing.Size(135, 30)
    $ok.DialogResult = 'OK'
    $ok.Anchor = 'Bottom, Right'
    _StylePrimaryButton $ok
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    [void](_AddStepBar $form 4)
    [void](_AddBrandHeader $form "Conclusions de l'informe" $null 44)

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }

    # Retornem els objectes triats (no nomes el text), per preservar
    # Title i Body per al desat de sessio i l'emissio al document.
    $selected = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $checks.Count; $i++) {
        if ($checks[$i].Checked) { [void]$selected.Add($checks[$i].Tag) }
    }
    return [pscustomobject]@{ Nav='next'; Data=(,$selected.ToArray()) }
}
