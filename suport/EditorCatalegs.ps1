#requires -Version 5.1
<#
.SYNOPSIS
  Editor visual dels ESTRUCTURALS (format estandard unic) des del programa.

.DESCRIPTION
  Una sola finestra "Editar catalegs" que edita QUALSEVOL ESTRUCTURAL en JSON
  (REQ1, TERMINI, 0 CONCLUSIONS, ACT_EXTR_REQ, ACT_EXTR_FAV) amb la MATEIXA
  interficie i el MATEIX vocabulari, perque tots comparteixen el mateix format:
    { tipus, familia, intro[], nodes[{tipus,titol,clau?,cos[],fills[]}] }
    <paragraf> = { runs:[{t,b,i}], url }

  - Un desplegable tria el document.
  - Un arbre (TreeView) mostra l'estructura (fills imbricats de veritat) + una
    entrada especial per a la "introduccio" (cos fix de TERMINI / capcalera de
    les conclusions).
  - A la dreta s'edita el titol, el TIPUS (mateixos noms a tots els catalegs;
    SEMPRE canviable quan hi ha mes d'una opcio) i el COS del node amb
    negreta/cursiva REALS (RichTextBox), i botons per inserir [CAMP:]/[OPCIO:] i
    enllacos. A ACT_EXTR es mostra tambe la CLAU funcional ([[KEY]]) en un camp
    a part i BLOQUEJAT (no s'edita mai des d'aqui).
  - Es poden afegir, eliminar i moure nodes. En desar s'escriu el JSON (amb
    validacio + copia .bak de seguretat).

  Separacio: les funcions PURES (conversio model<->JSON, runs<->segments) son
  testejables en headless; la interficie (WinForms) nomes s'executa a Windows.
  L'editor nomes ESCRIU JSON; la generacio (lectors ja validats) no en depen.
#>

# ===========================================================================
# FUNCIONS PURES: model editable <-> JSON
# ===========================================================================
# Model editable = hashtables/ArrayLists mutables (facils d'editar i de tornar a
# serialitzar amb ConvertTo-Json). Node = @{tipus;titol;clau;cos;fills};
# paragraf = @{runs=@(@{t;b;i});url}. 'clau' nomes te valor a ACT_EXTR (la [[KEY]]
# funcional); als altres catalegs es queda buida i no s'escriu al JSON.

function _Ed_ParasFromJson($paras) {
    $out = New-Object System.Collections.ArrayList
    foreach ($p in @($paras)) {
        $runs = New-Object System.Collections.ArrayList
        foreach ($r in @($p.runs)) {
            [void]$runs.Add(@{ t = [string]$r.t; b = [bool]$r.b; i = [bool]$r.i })
        }
        if ($runs.Count -eq 0) { [void]$runs.Add(@{ t = ''; b = $false; i = $false }) }
        [void]$out.Add(@{ runs = $runs; url = [bool]$p.url })
    }
    return ,$out
}

function _Ed_NodesFromJson($nodes) {
    $out = New-Object System.Collections.ArrayList
    foreach ($n in @($nodes)) {
        [void]$out.Add(@{
            tipus = [string]$n.tipus
            titol = [string]$n.titol
            clau  = [string]$n.clau
            cos   = (_Ed_ParasFromJson $n.cos)
            fills = (_Ed_NodesFromJson $n.fills)
        })
    }
    return ,$out
}

function _Ed_JsonToModel($o) {
    return @{
        tipus   = [string]$o.tipus
        familia = [string]$o.familia
        intro   = (_Ed_ParasFromJson $o.intro)
        nodes   = (_Ed_NodesFromJson $o.nodes)
    }
}

function _Ed_ParasToJson($paras) {
    $arr = @()
    foreach ($p in @($paras)) {
        $runs = @()
        foreach ($r in @($p.runs)) {
            $ro = [ordered]@{ t = [string]$r.t }
            if ($r.b) { $ro.b = $true }
            if ($r.i) { $ro.i = $true }
            $runs += ,$ro
        }
        if ($runs.Count -eq 0) { $runs = @([ordered]@{ t = '' }) }
        $po = [ordered]@{ runs = $runs }
        if ($p.url) { $po.url = $true }
        $arr += ,$po
    }
    return ,$arr
}

function _Ed_NodesToJson($nodes) {
    $arr = @()
    foreach ($n in @($nodes)) {
        # Ordre de claus identic al dels JSON generats: tipus, titol, [clau], cos,
        # fills. 'clau' nomes s'escriu si te valor (ACT_EXTR).
        $no = [ordered]@{
            tipus = [string]$n.tipus
            titol = [string]$n.titol
        }
        if (-not [string]::IsNullOrEmpty([string]$n.clau)) { $no.clau = [string]$n.clau }
        $no.cos   = (_Ed_ParasToJson $n.cos)
        $no.fills = (_Ed_NodesToJson $n.fills)
        $arr += ,$no
    }
    return ,$arr
}

function _Ed_ModelToJson($m) {
    $obj = [ordered]@{
        tipus   = [string]$m.tipus
        familia = [string]$m.familia
        intro   = (_Ed_ParasToJson $m.intro)
        nodes   = (_Ed_NodesToJson $m.nodes)
    }
    return ,$obj
}

# ===========================================================================
# FUNCIONS PURES: runs <-> "segments" (fragments amb Bold/Italic per al RTB)
# ===========================================================================
# Un "segment" = @{ Text; Bold; Italic }. Un "rich paragraf" = @{ Url; Segments }.
# Type-RichText tracta **/// com a spans NO solapats: un run es negreta O cursiva,
# mai totes dues. Per aixo, en serialitzar, si un segment porta les dues, es queda
# NOMES amb negreta (preserva la invariant i el round-trip amb la generacio).

function _Ed_SegmentsToRuns($segments) {
    $runs = New-Object System.Collections.ArrayList
    foreach ($s in @($segments)) {
        $txt = [string]$s.Text
        if ($txt.Length -eq 0) { continue }
        $b = [bool]$s.Bold
        $i = [bool]$s.Italic
        if ($b -and $i) { $i = $false }   # sense solapament: negreta guanya
        if ($runs.Count -gt 0) {
            $last = $runs[$runs.Count - 1]
            if ([bool]$last.b -eq $b -and [bool]$last.i -eq $i) {
                $last.t = [string]$last.t + $txt
                continue
            }
        }
        [void]$runs.Add(@{ t = $txt; b = $b; i = $i })
    }
    if ($runs.Count -eq 0) { [void]$runs.Add(@{ t = ''; b = $false; i = $false }) }
    return ,$runs
}

function _Ed_CosToRich($cosParas) {
    $out = New-Object System.Collections.ArrayList
    foreach ($p in @($cosParas)) {
        $segs = New-Object System.Collections.ArrayList
        foreach ($r in @($p.runs)) {
            [void]$segs.Add(@{ Text = [string]$r.t; Bold = [bool]$r.b; Italic = [bool]$r.i })
        }
        [void]$out.Add(@{ Url = [bool]$p.url; Segments = $segs })
    }
    return ,$out
}

function _Ed_RichToCos($richParas) {
    $out = New-Object System.Collections.ArrayList
    foreach ($rp in @($richParas)) {
        # Salta els paragrafs totalment buits (com feia el convertidor original).
        $joined = ''
        foreach ($s in @($rp.Segments)) { $joined += [string]$s.Text }
        if ([string]::IsNullOrWhiteSpace($joined) -and -not $rp.Url) { continue }
        if ([string]::IsNullOrWhiteSpace($joined) -and $rp.Url) { continue }
        [void]$out.Add(@{ runs = (_Ed_SegmentsToRuns $rp.Segments); url = [bool]$rp.Url })
    }
    return ,$out
}

# ===========================================================================
# FUNCIONS PURES: semantica del TIPUS (vocabulari unificat) per familia i pare
# ===========================================================================
# Vocabulari unic a tots els catalegs. Els tipus valids d'un node depenen de la
# FAMILIA i del TIPUS DEL PARE (arrel = pare buit ''); aixi la imbricacio es real
# (una subseccio conte items, un item conte subitems...). El lector torna a mapar
# cada tipus al Kind/Style intern, de manera que la generacio no canvia.
function _Ed_TipusOptions([string]$familia, [string]$parentTipus) {
    switch ($familia) {
        'cataleg' {
            if ([string]::IsNullOrEmpty($parentTipus)) { return @('seccio') }
            switch ($parentTipus) {
                'seccio'    { return @('item', 'subseccio', 'text') }
                'subseccio' { return @('item') }
                'item'      { return @('subitem') }
            }
            return @('item')
        }
        'conclusions' {
            if ([string]::IsNullOrEmpty($parentTipus)) { return @('seccio', 'sempre') }
            return @('item')
        }
        'actextr' {
            if ([string]::IsNullOrEmpty($parentTipus)) { return @('seccio') }
            # Sota una seccio (o dins d'un bloc) tots els estils de bloc son valids.
            return @('item', 'subitem', 'text', 'nota', 'etiqueta', 'capcalera', 'paragraf')
        }
    }
    return @('')
}

function _Ed_DefaultTipus([string]$familia, [string]$parentTipus) {
    return @(_Ed_TipusOptions $familia $parentTipus)[0]
}

# Un node d'aquesta familia pot tenir fills? (nomes on el lector els llegeix.)
function _Ed_CanAddChild([string]$familia, $node) {
    switch ($familia) {
        'cataleg'     { return ([string]$node.tipus -in @('seccio', 'subseccio', 'item')) }
        'conclusions' { return ([string]$node.tipus -eq 'seccio') }
        'actextr'     { return ([string]$node.tipus -in @('seccio', 'item')) }
    }
    return $false
}

# El tipus d'un fill nou, segons familia i tipus del pare (sempre dins de
# _Ed_TipusOptions del pare, perque el combo el pugui mostrar).
function _Ed_ChildTipus([string]$familia, [string]$parentTipus) {
    switch ($familia) {
        'cataleg'     { if ($parentTipus -eq 'seccio' -or $parentTipus -eq 'subseccio') { return 'item' } else { return 'subitem' } }
        'conclusions' { return 'item' }
        'actextr'     { if ($parentTipus -eq 'seccio') { return 'item' } else { return 'subitem' } }
    }
    return (_Ed_DefaultTipus $familia $parentTipus)
}

# Text curt per a l'etiqueta d'un node a l'arbre.
function _Ed_NodeLabel($node) {
    $t = [string]$node.titol
    if ([string]::IsNullOrWhiteSpace($t) -and @($node.cos).Count -gt 0) {
        $t = _RunsToMarkup (@($node.cos)[0]).runs
    }
    $t = ($t -replace '\s+', ' ').Trim()
    if ($t.Length -gt 64) { $t = $t.Substring(0, 64) + [char]0x2026 }
    return ('[' + [string]$node.tipus + '] ' + $t)
}

# Llista dels ESTRUCTURALS editables (tots els *.json de la carpeta), amb nom
# amic per als coneguts. Retorna @(@{ Key; Label; Path }).
function _Ed_DocList {
    $known = @{
        'REQ1'         = 'REQ1 ' + [char]0x2014 + ' Requeriment (cat' + [char]0x00E0 + 'leg)'
        'TERMINI'      = 'TERMINI ' + [char]0x2014 + ' Ampliaci' + [char]0x00F3 + ' termini'
        '0 CONCLUSIONS'= '0 CONCLUSIONS ' + [char]0x2014 + ' Conclusions'
        'ACT_EXTR_REQ' = 'ACT_EXTR_REQ ' + [char]0x2014 + ' Requeriment act. extraordin' + [char]0x00E0 + 'ria'
        'ACT_EXTR_FAV' = 'ACT_EXTR_FAV ' + [char]0x2014 + ' Informe favorable act. extraordin' + [char]0x00E0 + 'ria'
    }
    $out = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $EstructuralsDir)) { return $out }
    foreach ($f in (Get-ChildItem -LiteralPath $EstructuralsDir -Filter '*.json' | Sort-Object Name)) {
        $key = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $label = if ($known.ContainsKey($key)) { $known[$key] } else { $key }
        [void]$out.Add(@{ Key = $key; Label = $label; Path = $f.FullName })
    }
    return $out
}

# ===========================================================================
# INTERFICIE (WinForms) - nomes s'executa a Windows
# ===========================================================================

# Llegeix el RichTextBox i el converteix a rich paragraphs (@{Url;Segments}).
# Recorre caracter a caracter (els cossos son curts) per llegir negreta/cursiva.
function _Ed_ReadRtbToRich($rtb) {
    $result = New-Object System.Collections.ArrayList
    $text = [string]$rtb.Text
    $n = $rtb.TextLength
    $selStart = $rtb.SelectionStart
    $selLen = $rtb.SelectionLength

    # Recull per caracter (c, bold, italic), separant paragrafs pel salt de linia.
    $paraChars = New-Object System.Collections.ArrayList
    $cur = New-Object System.Collections.ArrayList
    for ($k = 0; $k -lt $n; $k++) {
        $c = $text[$k]
        if ($c -eq [char]10) { [void]$paraChars.Add($cur); $cur = New-Object System.Collections.ArrayList; continue }
        if ($c -eq [char]13) { continue }
        $rtb.Select($k, 1)
        $f = $rtb.SelectionFont
        $b = if ($null -ne $f) { $f.Bold } else { $false }
        $it = if ($null -ne $f) { $f.Italic } else { $false }
        [void]$cur.Add(@{ c = $c; b = $b; i = $it })
    }
    [void]$paraChars.Add($cur)
    $rtb.Select($selStart, $selLen)

    foreach ($cl in $paraChars) {
        $chars = @($cl)
        $url = $false
        $s = -join ($chars | ForEach-Object { $_.c })
        if ($s.StartsWith('[[URL]] ')) {
            $url = $true
            if ($chars.Count -gt 8) { $chars = $chars[8..($chars.Count - 1)] } else { $chars = @() }
        }
        $segs = New-Object System.Collections.ArrayList
        foreach ($ch in $chars) {
            if ($segs.Count -gt 0) {
                $last = $segs[$segs.Count - 1]
                if ([bool]$last.Bold -eq [bool]$ch.b -and [bool]$last.Italic -eq [bool]$ch.i) {
                    $last.Text = [string]$last.Text + [string]$ch.c
                    continue
                }
            }
            [void]$segs.Add(@{ Text = [string]$ch.c; Bold = [bool]$ch.b; Italic = [bool]$ch.i })
        }
        [void]$result.Add(@{ Url = $url; Segments = $segs })
    }
    return ,$result
}

# Renderitza rich paragraphs al RichTextBox (negreta/cursiva reals; el prefix
# [[URL]] dels enllacos es mostra en gris).
function _Ed_RenderRichToRtb($rtb, $richParas, $baseFont) {
    $rtb.Clear()
    $grey = [System.Drawing.Color]::FromArgb(120, 128, 138)
    $ink = [System.Drawing.Color]::FromArgb(29, 39, 51)
    $first = $true
    foreach ($rp in @($richParas)) {
        if (-not $first) { $rtb.AppendText("`n") }
        $first = $false
        if ($rp.Url) {
            $st = $rtb.TextLength
            $rtb.AppendText('[[URL]] ')
            $rtb.Select($st, 8)
            $rtb.SelectionColor = $grey
            $rtb.SelectionFont = New-Object System.Drawing.Font($baseFont.FontFamily, $baseFont.Size, [System.Drawing.FontStyle]::Regular)
        }
        foreach ($seg in @($rp.Segments)) {
            $txt = [string]$seg.Text
            if ($txt.Length -eq 0) { continue }
            $st = $rtb.TextLength
            $rtb.AppendText($txt)
            $rtb.Select($st, $txt.Length)
            $style = [System.Drawing.FontStyle]::Regular
            if ($seg.Bold) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
            if ($seg.Italic) { $style = $style -bor [System.Drawing.FontStyle]::Italic }
            $rtb.SelectionColor = $ink
            $rtb.SelectionFont = New-Object System.Drawing.Font($baseFont.FontFamily, $baseFont.Size, $style)
        }
    }
    $rtb.Select(0, 0)
    $rtb.SelectionColor = $ink
    $rtb.SelectionFont = $baseFont
}

# Aplica negreta o cursiva a la seleccio del RTB (mutualment excloents, per
# respectar la invariant de Type-RichText).
function _Ed_ToggleStyle($rtb, [string]$which, $baseFont) {
    if ($rtb.SelectionLength -le 0) { return }
    $f = $rtb.SelectionFont
    if ($which -eq 'bold') {
        $makeOn = -not ($null -ne $f -and $f.Bold)
        $style = if ($makeOn) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    } else {
        $makeOn = -not ($null -ne $f -and $f.Italic)
        $style = if ($makeOn) { [System.Drawing.FontStyle]::Italic } else { [System.Drawing.FontStyle]::Regular }
    }
    $rtb.SelectionFont = New-Object System.Drawing.Font($baseFont.FontFamily, $baseFont.Size, $style)
}

# Insereix text pla al RTB en la posicio del cursor (estil normal).
function _Ed_InsertPlain($rtb, [string]$text, $baseFont) {
    $st = $rtb.SelectionStart
    $rtb.SelectedText = $text
    $rtb.Select($st, $text.Length)
    $rtb.SelectionFont = New-Object System.Drawing.Font($baseFont.FontFamily, $baseFont.Size, [System.Drawing.FontStyle]::Regular)
    $rtb.Select($st + $text.Length, 0)
    $rtb.Focus()
}

# ---- Arbre ----------------------------------------------------------------
# El tag de cada TreeNode guarda: el node de model, la llista de germans (per
# moure/eliminar) i el NODE PARE ($null a l'arrel) per calcular els tipus valids.
function _Ed_AddTreeNode($parentNodes, $node, $siblings, $parentNode) {
    $tn = New-Object System.Windows.Forms.TreeNode((_Ed_NodeLabel $node))
    $tn.Tag = @{ Kind = 'node'; Node = $node; Parent = $siblings; ParentNode = $parentNode }
    [void]$parentNodes.Add($tn)
    foreach ($ch in @($node.fills)) { _Ed_AddTreeNode $tn.Nodes $ch $node.fills $node }
    $tn.Expand()
    return $tn
}

function _Ed_BuildTree($state) {
    $tree = $state.Tree
    $m = $state.Model
    $state.Busy = $true
    $tree.BeginUpdate()
    $tree.Nodes.Clear()
    $introLabel = switch ($m.familia) {
        'conclusions' { [char]0x00B7 + ' Cap' + [char]0x00E7 + 'alera' }
        default { if (@($m.nodes).Count -eq 0) { [char]0x00B7 + ' Cos fix' } else { [char]0x00B7 + ' Introducci' + [char]0x00F3 } }
    }
    $tnIntro = New-Object System.Windows.Forms.TreeNode($introLabel)
    $tnIntro.Tag = @{ Kind = 'intro' }
    [void]$tree.Nodes.Add($tnIntro)
    foreach ($n in @($m.nodes)) { [void](_Ed_AddTreeNode $tree.Nodes $n $m.nodes $null) }
    $tree.EndUpdate()
    $state.Busy = $false
}

# Selecciona el TreeNode que embolcalla el node de model $target (per referencia).
function _Ed_FindTreeNode($nodes, $target) {
    foreach ($tn in $nodes) {
        if ($tn.Tag.Kind -eq 'node' -and [object]::ReferenceEquals($tn.Tag.Node, $target)) { return $tn }
        $r = _Ed_FindTreeNode $tn.Nodes $target
        if ($null -ne $r) { return $r }
    }
    return $null
}

function _Ed_SelectModelNode($state, $target) {
    $tn = _Ed_FindTreeNode $state.Tree.Nodes $target
    if ($null -ne $tn) {
        $state.Busy = $true
        $state.Tree.SelectedNode = $tn
        $state.Busy = $false
        $state.Bound = $tn.Tag
        _Ed_LoadEditor $state
    }
}

# ---- Carrega/desa els controls de la dreta <-> node lligat ($state.Bound) ---
function _Ed_LoadEditor($state) {
    $state.Busy = $true
    $tag = $state.Bound
    $isActextr = ([string]$state.Model.familia -eq 'actextr')
    if ($null -eq $tag) {
        $state.TitolBox.Text = ''
        $state.TitolBox.Enabled = $false
        $state.TipusCombo.Items.Clear()
        $state.TipusCombo.Enabled = $false
        $state.ClauLabel.Visible = $false
        $state.ClauBox.Visible = $false
        $state.ClauBox.Text = ''
        $state.Rtb.Clear()
        $state.Rtb.Enabled = $false
        $state.Busy = $false
        return
    }
    if ($tag.Kind -eq 'intro') {
        $state.TitolBox.Text = ''
        $state.TitolBox.Enabled = $false
        $state.TipusCombo.Items.Clear()
        $state.TipusCombo.Enabled = $false
        $state.ClauLabel.Visible = $false
        $state.ClauBox.Visible = $false
        $state.ClauBox.Text = ''
        $state.Rtb.Enabled = $true
        _Ed_RenderRichToRtb $state.Rtb (_Ed_CosToRich $state.Model.intro) $state.RtbFont
    } else {
        $node = $tag.Node
        $state.TitolBox.Enabled = $true
        $state.TitolBox.Text = [string]$node.titol
        $parentTipus = if ($null -ne $tag.ParentNode) { [string]$tag.ParentNode.tipus } else { '' }
        $opts = @(_Ed_TipusOptions $state.Model.familia $parentTipus)
        $state.TipusCombo.Items.Clear()
        foreach ($o in $opts) { [void]$state.TipusCombo.Items.Add($o) }
        $idx = $state.TipusCombo.Items.IndexOf([string]$node.tipus)
        if ($idx -lt 0 -and $state.TipusCombo.Items.Count -gt 0) { $idx = 0 }
        if ($idx -ge 0) { $state.TipusCombo.SelectedIndex = $idx }
        # SEMPRE canviable quan hi ha mes d'una opcio.
        $state.TipusCombo.Enabled = ($opts.Count -gt 1)
        # Camp CLAU: nomes a ACT_EXTR, i BLOQUEJAT (mostra la [[KEY]] funcional).
        $state.ClauLabel.Visible = $isActextr
        $state.ClauBox.Visible = $isActextr
        $state.ClauBox.Text = if ($isActextr) { [string]$node.clau } else { '' }
        $state.Rtb.Enabled = $true
        _Ed_RenderRichToRtb $state.Rtb (_Ed_CosToRich $node.cos) $state.RtbFont
    }
    $state.Busy = $false
}

function _Ed_FlushEditor($state) {
    $tag = $state.Bound
    if ($null -eq $tag) { return }
    if ($tag.Kind -eq 'intro') {
        $state.Model.intro = _Ed_RichToCos (_Ed_ReadRtbToRich $state.Rtb)
        return
    }
    $node = $tag.Node
    $node.titol = [string]$state.TitolBox.Text
    if ($state.TipusCombo.Enabled -and $null -ne $state.TipusCombo.SelectedItem) {
        $node.tipus = [string]$state.TipusCombo.SelectedItem
    }
    # 'clau' es nomes lectura: no es toca mai des de l'editor.
    $node.cos = _Ed_RichToCos (_Ed_ReadRtbToRich $state.Rtb)
}

# Refresca EN VIU l'etiqueta del node seleccionat a l'arbre (mentre s'edita, el
# node seleccionat ES el que s'esta editant, aixi que es segur).
function _Ed_RelabelSelected($state) {
    $tag = $state.Bound
    if ($null -eq $tag -or $tag.Kind -ne 'node') { return }
    if ($null -eq $state.Tree.SelectedNode) { return }
    $tipus = if ($null -ne $state.TipusCombo.SelectedItem) { [string]$state.TipusCombo.SelectedItem } else { [string]$tag.Node.tipus }
    $t = ([string]$state.TitolBox.Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt 64) { $t = $t.Substring(0, 64) + [char]0x2026 }
    $state.Tree.SelectedNode.Text = ('[' + $tipus + '] ' + $t)
}

function _Ed_OnTreeSelect($state) {
    if ($state.Busy) { return }
    if ([object]::ReferenceEquals($state.Bound, $state.Tree.SelectedNode.Tag)) { return }
    _Ed_FlushEditor $state
    $state.Bound = $state.Tree.SelectedNode.Tag
    _Ed_LoadEditor $state
}

# ---- Operacions d'estructura ----------------------------------------------
function _Ed_NewNode([string]$tipus) {
    return @{
        tipus = $tipus; titol = ''; clau = ''
        cos = (New-Object System.Collections.ArrayList)
        fills = (New-Object System.Collections.ArrayList)
    }
}

function _Ed_AddSibling($state) {
    _Ed_FlushEditor $state
    $tag = $state.Bound
    if ($null -eq $tag -or $tag.Kind -eq 'intro') {
        # Sense node (o intro): afegeix un node d'arrel al final.
        $tipus = _Ed_DefaultTipus $state.Model.familia ''
        $new = _Ed_NewNode $tipus
        [void]$state.Model.nodes.Add($new)
    } else {
        $sib = $tag.Parent
        $node = $tag.Node
        $idx = $sib.IndexOf($node)
        # El germa nou comparteix el pare del node actual: mateix conjunt de tipus.
        $parentTipus = if ($null -ne $tag.ParentNode) { [string]$tag.ParentNode.tipus } else { '' }
        $opts = @(_Ed_TipusOptions $state.Model.familia $parentTipus)
        $tipus = if ($opts -contains [string]$node.tipus) { [string]$node.tipus } else { $opts[0] }
        $new = _Ed_NewNode $tipus
        if ($idx -ge 0) { $sib.Insert($idx + 1, $new) } else { [void]$sib.Add($new) }
    }
    $state.Dirty = $true
    _Ed_BuildTree $state
    _Ed_SelectModelNode $state $new
}

function _Ed_AddChild($state) {
    _Ed_FlushEditor $state
    $tag = $state.Bound
    if ($null -eq $tag -or $tag.Kind -eq 'intro') { return }
    $node = $tag.Node
    if (-not (_Ed_CanAddChild $state.Model.familia $node)) {
        [System.Windows.Forms.MessageBox]::Show('Aquest tipus de node no pot tenir fills.', 'Editar catalegs', 'OK', 'Information') | Out-Null
        return
    }
    $tipus = _Ed_ChildTipus $state.Model.familia ([string]$node.tipus)
    $new = _Ed_NewNode $tipus
    [void]$node.fills.Add($new)
    $state.Dirty = $true
    _Ed_BuildTree $state
    _Ed_SelectModelNode $state $new
}

function _Ed_DeleteNode($state) {
    $tag = $state.Bound
    if ($null -eq $tag -or $tag.Kind -eq 'intro') { return }
    $node = $tag.Node
    $r = [System.Windows.Forms.MessageBox]::Show('Segur que vols eliminar aquest node i tot el que conté?', 'Eliminar node', 'YesNo', 'Warning')
    if ($r -ne 'Yes') { return }
    [void]$tag.Parent.Remove($node)
    $state.Bound = $null
    $state.Dirty = $true
    _Ed_BuildTree $state
    _Ed_LoadEditor $state
}

function _Ed_MoveNode($state, [int]$delta) {
    _Ed_FlushEditor $state
    $tag = $state.Bound
    if ($null -eq $tag -or $tag.Kind -eq 'intro') { return }
    $node = $tag.Node
    $sib = $tag.Parent
    $idx = $sib.IndexOf($node)
    $new = $idx + $delta
    if ($idx -lt 0 -or $new -lt 0 -or $new -ge $sib.Count) { return }
    $sib.RemoveAt($idx)
    $sib.Insert($new, $node)
    $state.Dirty = $true
    _Ed_BuildTree $state
    _Ed_SelectModelNode $state $node
}

# ---- Carrega / desa document ----------------------------------------------
function _Ed_LoadDoc($state, $doc) {
    $o = Get-Content -LiteralPath $doc.Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $state.Model = _Ed_JsonToModel $o
    $state.CurrentDoc = $doc
    $state.Bound = $null
    $state.Dirty = $false
    _Ed_BuildTree $state
    # Selecciona el primer node real (o la introduccio si no n'hi ha).
    if (@($state.Model.nodes).Count -gt 0) {
        _Ed_SelectModelNode $state (@($state.Model.nodes)[0])
    } else {
        $state.Busy = $true
        $state.Tree.SelectedNode = $state.Tree.Nodes[0]
        $state.Busy = $false
        $state.Bound = $state.Tree.Nodes[0].Tag
        _Ed_LoadEditor $state
    }
}

function _Ed_SaveDoc($state) {
    _Ed_FlushEditor $state
    $doc = $state.CurrentDoc
    if ($null -eq $doc) { return }
    $obj = _Ed_ModelToJson $state.Model
    $json = $obj | ConvertTo-Json -Depth 40

    # Validacio: escriu a temporal i comprova que el lector corresponent no peta.
    $tmp = [System.IO.Path]::GetTempFileName()
    try {
        [System.IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding($false)))
        switch ([string]$state.Model.familia) {
            'cataleg'     { [void](Read-CatalegJson $tmp) }
            'conclusions' { [void](Read-ConclusionsJson $tmp '') }
            'actextr'     { [void](Read-ActExtrRecordsJson $tmp) }
        }
    } catch {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show("No s'ha desat: el document no es valid.`n`n$($_.Exception.Message)", 'Editar catalegs', 'OK', 'Error') | Out-Null
        return
    }
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue

    try {
        if (Test-Path -LiteralPath $doc.Path) { Copy-Item -LiteralPath $doc.Path -Destination ($doc.Path + '.bak') -Force }
        [System.IO.File]::WriteAllText($doc.Path, $json, (New-Object System.Text.UTF8Encoding($false)))
        $state.Dirty = $false
        [System.Windows.Forms.MessageBox]::Show('Catàleg desat correctament.', 'Editar catalegs', 'OK', 'Information') | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error en desar:`n$($_.Exception.Message)", 'Editar catalegs', 'OK', 'Error') | Out-Null
    }
}

# Pregunta si cal desar quan hi ha canvis pendents. Retorna $true si es pot
# continuar (desat o descartat), $false si l'usuari cancel-la.
function _Ed_ConfirmDiscard($state) {
    if (-not $state.Dirty) { return $true }
    $r = [System.Windows.Forms.MessageBox]::Show('Hi ha canvis sense desar. Vols desar-los?', 'Editar catalegs', 'YesNoCancel', 'Warning')
    if ($r -eq 'Cancel') { return $false }
    if ($r -eq 'Yes') { _Ed_SaveDoc $state }
    return $true
}

# ---- Finestra principal ---------------------------------------------------
# $focusDoc: clau del document a obrir de bon principi (REQ1, TERMINI,
# ACT_EXTR[_REQ/_FAV], '0 CONCLUSIONS'...). ACT_EXTR sol -> ACT_EXTR_REQ.
function Show-CatalegEditor([string]$focusDoc = '') {
    $docs = @(_Ed_DocList)
    if ($docs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No hi ha cap catàleg JSON a ESTRUCTURALS.', 'Editar catalegs', 'OK', 'Information') | Out-Null
        return
    }

    $form = _NewForm
    $form.Text = 'Editar catalegs'
    $form.ClientSize = New-Object System.Drawing.Size(970, 660)
    $form.MinimumSize = New-Object System.Drawing.Size(836, 700)

    $baseRtbFont = New-Object System.Drawing.Font('Segoe UI', 10.5, [System.Drawing.FontStyle]::Regular)

    $state = @{
        Model = $null; CurrentDoc = $null; Bound = $null; Dirty = $false; Busy = $false
        Tree = $null; TitolBox = $null; TipusCombo = $null; ClauLabel = $null; ClauBox = $null
        Rtb = $null; RtbFont = $baseRtbFont
    }

    $yTop = 66

    # Selector de document.
    $lblDoc = New-Object System.Windows.Forms.Label
    $lblDoc.Text = 'Document:'
    $lblDoc.Location = New-Object System.Drawing.Point(16, ($yTop + 3))
    $lblDoc.AutoSize = $true
    [void]$form.Controls.Add($lblDoc)

    $cbDoc = New-Object System.Windows.Forms.ComboBox
    $cbDoc.DropDownStyle = 'DropDownList'
    $cbDoc.Location = New-Object System.Drawing.Point(90, $yTop)
    $cbDoc.Size = New-Object System.Drawing.Size(360, 26)
    foreach ($d in $docs) { [void]$cbDoc.Items.Add($d.Label) }
    [void]$form.Controls.Add($cbDoc)

    # Arbre.
    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Location = New-Object System.Drawing.Point(16, ($yTop + 36))
    $tree.Size = New-Object System.Drawing.Size(360, 420)
    $tree.Anchor = 'Top, Bottom, Left'
    $tree.HideSelection = $false
    $tree.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
    [void]$form.Controls.Add($tree)
    $state.Tree = $tree

    # Botons d'estructura, sota l'arbre.
    $yStruct = $yTop + 36 + 420 + 8
    $mkStructBtn = {
        param($text, $x, $y, $w)
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $text
        $b.Location = New-Object System.Drawing.Point($x, $y)
        $b.Size = New-Object System.Drawing.Size($w, 28)
        $b.Anchor = 'Bottom, Left'
        _StyleSecondaryButton $b
        [void]$form.Controls.Add($b)
        return $b
    }
    $btnAdd    = & $mkStructBtn ('+ Node')  16  $yStruct 78
    $btnAddSub = & $mkStructBtn ('+ Fill')  98  $yStruct 78
    $btnDel    = & $mkStructBtn 'Eliminar'  180 $yStruct 90
    $btnUp     = & $mkStructBtn ([char]0x2191 + ' Pujar')  16 ($yStruct + 34) 120
    $btnDown   = & $mkStructBtn ([char]0x2193 + ' Baixar') 140 ($yStruct + 34) 120

    # Editor de la dreta: titol.
    $xR = 396
    $lblTit = New-Object System.Windows.Forms.Label
    $lblTit.Text = ('T' + [char]0x00ED + 'tol / etiqueta:')
    $lblTit.Location = New-Object System.Drawing.Point($xR, ($yTop + 3))
    $lblTit.AutoSize = $true
    [void]$form.Controls.Add($lblTit)

    $tbTit = New-Object System.Windows.Forms.TextBox
    $tbTit.Location = New-Object System.Drawing.Point($xR, ($yTop + 24))
    $tbTit.Size = New-Object System.Drawing.Size(558, 26)
    $tbTit.Anchor = 'Top, Left, Right'
    [void]$form.Controls.Add($tbTit)
    $state.TitolBox = $tbTit

    # Tipus (vocabulari unificat, sempre canviable quan hi ha >1 opcio).
    $lblMar = New-Object System.Windows.Forms.Label
    $lblMar.Text = 'Tipus:'
    $lblMar.Location = New-Object System.Drawing.Point($xR, ($yTop + 58))
    $lblMar.AutoSize = $true
    [void]$form.Controls.Add($lblMar)

    $cbMar = New-Object System.Windows.Forms.ComboBox
    $cbMar.DropDownStyle = 'DropDownList'
    $cbMar.Location = New-Object System.Drawing.Point(($xR + 52), ($yTop + 55))
    $cbMar.Size = New-Object System.Drawing.Size(160, 26)
    [void]$form.Controls.Add($cbMar)
    $state.TipusCombo = $cbMar

    # Clau (nomes ACT_EXTR): la [[KEY]] funcional, BLOQUEJADA (nomes lectura).
    $lblClau = New-Object System.Windows.Forms.Label
    $lblClau.Text = 'Clau:'
    $lblClau.Location = New-Object System.Drawing.Point(($xR + 224), ($yTop + 58))
    $lblClau.AutoSize = $true
    $lblClau.Visible = $false
    [void]$form.Controls.Add($lblClau)
    $state.ClauLabel = $lblClau

    $tbClau = New-Object System.Windows.Forms.TextBox
    $tbClau.Location = New-Object System.Drawing.Point(($xR + 262), ($yTop + 55))
    $tbClau.Size = New-Object System.Drawing.Size(296, 26)
    $tbClau.Anchor = 'Top, Left, Right'
    $tbClau.ReadOnly = $true
    $tbClau.TabStop = $false
    $tbClau.BackColor = [System.Drawing.Color]::FromArgb(238, 241, 245)
    $tbClau.ForeColor = [System.Drawing.Color]::FromArgb(90, 99, 110)
    $tbClau.Visible = $false
    [void]$form.Controls.Add($tbClau)
    $state.ClauBox = $tbClau

    # Barra d'eines del cos.
    $yTool = $yTop + 90
    $mkToolBtn = {
        param($text, $x, $w, $bold)
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $text
        $b.Location = New-Object System.Drawing.Point($x, $yTool)
        $b.Size = New-Object System.Drawing.Size($w, 28)
        $b.Anchor = 'Top, Left'
        _StyleSecondaryButton $b
        if ($bold) { $b.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold) }
        [void]$form.Controls.Add($b)
        return $b
    }
    $btnBold   = & $mkToolBtn 'N'        $xR 40 $true
    $btnItal   = & $mkToolBtn 'C'        ($xR + 46) 40 $false
    $btnItal.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Italic)
    $btnCamp   = & $mkToolBtn '[CAMP]'   ($xR + 96) 78 $false
    $btnOpcio  = & $mkToolBtn '[OPCIO]'  ($xR + 180) 78 $false
    $btnLink   = & $mkToolBtn ([System.Char]::ConvertFromUtf32(0x1F517) + ' Enlla' + [char]0x00E7) ($xR + 264) 100 $false

    # Cos (RichTextBox).
    $lblCos = New-Object System.Windows.Forms.Label
    $lblCos.Text = 'Cos (selecciona i prem N/C per negreta/cursiva):'
    $lblCos.Location = New-Object System.Drawing.Point($xR, ($yTool + 34))
    $lblCos.AutoSize = $true
    $lblCos.ForeColor = [System.Drawing.Color]::FromArgb(107, 116, 128)
    [void]$form.Controls.Add($lblCos)

    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.Location = New-Object System.Drawing.Point($xR, ($yTool + 56))
    $rtb.Size = New-Object System.Drawing.Size(558, 360)
    $rtb.Anchor = 'Top, Bottom, Left, Right'
    $rtb.Font = $baseRtbFont
    $rtb.AcceptsTab = $false
    $rtb.HideSelection = $false
    [void]$form.Controls.Add($rtb)
    $state.Rtb = $rtb

    # Barra inferior: Enrere / Desar.
    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = 'Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(16, 620)
    $btnBack.Size = New-Object System.Drawing.Size(110, 30)
    $btnBack.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $btnBack
    [void]$form.Controls.Add($btnBack)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Desar'
    $btnSave.Location = New-Object System.Drawing.Point(834, 620)
    $btnSave.Size = New-Object System.Drawing.Size(120, 30)
    $btnSave.Anchor = 'Bottom, Right'
    _StylePrimaryButton $btnSave
    [void]$form.Controls.Add($btnSave)

    # Capcalera de marca (Dock=Top): s'afegeix DESPRES dels controls absoluts.
    [void](_AddBrandHeader $form ('Editar cat' + [char]0x00E0 + 'legs') ('Format estàndard únic ' + [char]0x00B7 + ' es desa al JSON'))

    # ---- Handlers (capturen nomes $state / controls per closure) ----------
    $tree.add_AfterSelect({ param($s, $e) _Ed_OnTreeSelect $state }.GetNewClosure())
    $btnBold.add_Click({ _Ed_ToggleStyle $state.Rtb 'bold' $state.RtbFont; $state.Rtb.Focus(); $state.Dirty = $true }.GetNewClosure())
    $btnItal.add_Click({ _Ed_ToggleStyle $state.Rtb 'italic' $state.RtbFont; $state.Rtb.Focus(); $state.Dirty = $true }.GetNewClosure())
    $btnCamp.add_Click({ _Ed_InsertPlain $state.Rtb '[CAMP: etiqueta]' $state.RtbFont; $state.Dirty = $true }.GetNewClosure())
    $btnOpcio.add_Click({ _Ed_InsertPlain $state.Rtb '[OPCIO: A | B | C]' $state.RtbFont; $state.Dirty = $true }.GetNewClosure())
    $btnLink.add_Click({
        # Afegeix una nova linia d'enllac al final del cos.
        if ($state.Rtb.TextLength -gt 0) { $state.Rtb.AppendText("`n") }
        _Ed_InsertPlain $state.Rtb '[[URL]] https://' $state.RtbFont
        $state.Dirty = $true
    }.GetNewClosure())
    $rtb.add_TextChanged({ if (-not $state.Busy) { $state.Dirty = $true } }.GetNewClosure())
    $tbTit.add_TextChanged({ if (-not $state.Busy) { $state.Dirty = $true; _Ed_RelabelSelected $state } }.GetNewClosure())
    $cbMar.add_SelectedIndexChanged({
        if ($state.Busy) { return }
        $state.Dirty = $true
        # Aplica el canvi de tipus al node lligat i refresca l'etiqueta de l'arbre.
        $tag = $state.Bound
        if ($null -ne $tag -and $tag.Kind -eq 'node' -and $null -ne $state.TipusCombo.SelectedItem) {
            $tag.Node.tipus = [string]$state.TipusCombo.SelectedItem
        }
        _Ed_RelabelSelected $state
    }.GetNewClosure())

    $btnAdd.add_Click({ _Ed_AddSibling $state }.GetNewClosure())
    $btnAddSub.add_Click({ _Ed_AddChild $state }.GetNewClosure())
    $btnDel.add_Click({ _Ed_DeleteNode $state }.GetNewClosure())
    $btnUp.add_Click({ _Ed_MoveNode $state -1 }.GetNewClosure())
    $btnDown.add_Click({ _Ed_MoveNode $state 1 }.GetNewClosure())
    $btnSave.add_Click({ _Ed_SaveDoc $state }.GetNewClosure())
    # Nomes tanca: el FormClosing ja demana si cal desar (evitem doble pregunta).
    $btnBack.add_Click({ $form.Close() }.GetNewClosure())

    $cbDoc.add_SelectedIndexChanged({
        if ($state.Busy) { return }
        $sel = $cbDoc.SelectedIndex
        if ($sel -lt 0) { return }
        if (-not (_Ed_ConfirmDiscard $state)) {
            # Torna a seleccionar el document actual sense recarregar.
            $state.Busy = $true
            for ($j = 0; $j -lt $docs.Count; $j++) { if ($docs[$j].Key -eq $state.CurrentDoc.Key) { $cbDoc.SelectedIndex = $j } }
            $state.Busy = $false
            return
        }
        _Ed_LoadDoc $state $docs[$sel]
    }.GetNewClosure())

    $form.add_FormClosing({
        param($s, $e)
        if (-not (_Ed_ConfirmDiscard $state)) { $e.Cancel = $true }
    }.GetNewClosure())

    # Document inicial: el que ve del xip (focusDoc); ACT_EXTR -> ACT_EXTR_REQ.
    $wantKey = [string]$focusDoc
    if ($wantKey -eq 'ACT_EXTR') { $wantKey = 'ACT_EXTR_REQ' }
    $startIdx = 0
    for ($j = 0; $j -lt $docs.Count; $j++) { if ($docs[$j].Key -eq $wantKey) { $startIdx = $j; break } }
    $state.Busy = $true
    $cbDoc.SelectedIndex = $startIdx
    $state.Busy = $false
    _Ed_LoadDoc $state $docs[$startIdx]

    [void]$form.ShowDialog()
    $form.Dispose()
}
