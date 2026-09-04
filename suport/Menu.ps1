#requires -Version 5.1
<#
.SYNOPSIS
  El menu principal del programa (Pas 1) i el segell d'ultima execucio.

.DESCRIPTION
  Select-Mode vivia dins de Seguiment.ps1 -el fitxer de l'eina "Informe de
  seguiment"-, i son 550 linies que no hi tenen res a veure: es LA PRIMERA
  PANTALLA del programa, la que crida Main (Wizard.ps1) i des d'on surten totes
  les eines. Que estava mal posat ho deia la propia suite: QUATRE dels sis
  guards que llegien Seguiment.ps1 eren guards del MENU (el titol acotat pel
  xip, el $result.Choice amb closure, el boto de la carpeta i l'ordre dels
  informes).

  Ve amb ell el SEGELL d'ultima execucio, que nomes serveix per pintar la data
  sota cada rajola.

  Nomes defineix; les dues variables de nivell superior
  ($Script:AccionsSenseSegell i $Script:SegellPropi) son llistes, no calculen res.
#>

# Pantalla inicial (Pas 1): un sol menu que fusiona la tria de MODE i la de
# CATALEG. Cada boto mostra un nom amic (negre) i, en GRIS, el nom del document
# d'ESTRUCTURALS entre parentesis perque no destaqui tant.
#
# Formata una marca de temps ISO per mostrar-la (petita) sota les rajoles del
# menu. '(mai)' si es buida o no es pot llegir. Funcio PURA (testejable).
function _FormatRunStamp([string]$iso) {
    if ([string]::IsNullOrWhiteSpace($iso)) { return '(mai)' }
    try { return ([datetime]::Parse($iso)).ToString('dd/MM/yy HH:mm') } catch { return '(mai)' }
}

# Llegeix una marca de temps d'"ultima execucio" d'un JSON d'estat i la formata.
function _LastRunText($jsonPath, $prop) {
    if ([string]::IsNullOrWhiteSpace($jsonPath) -or -not (Test-Path -LiteralPath $jsonPath -ErrorAction SilentlyContinue)) { return '(mai)' }
    try {
        $o = Read-JsonFile $jsonPath
        if ($null -ne $o -and $o.PSObject.Properties[$prop] -and -not [string]::IsNullOrWhiteSpace([string]$o.$prop)) {
            return (_FormatRunStamp ([string]$o.$prop))
        }
    } catch { }
    return '(mai)'
}

# ----------------------------------------------------------------------------
# SEGELL D'ULTIMA EXECUCIO DE LES EINES
# ----------------------------------------------------------------------------
# Sota cada rajola del menu hi surt quan es va fer servir aquella eina per
# ultima vegada. Un SOL registre per a totes, indexat per ACCIO:
#
#     local\base-dades-activitats\eines-state.json   ->  { "<accio>": "<ISO>" }
#
# Abans cada segell tenia el seu fitxer i el seu nom de propietat i el menu els
# llegia un per un; amb onze rajoles aixo seria pura duplicitat. Es marca en UN
# SOL LLOC: al despatxador de Main (Wizard.ps1), quan l'eina torna. Per tant la
# data vol dir "l'ultima vegada que has obert i tancat aquesta eina".
#
# Accions que NO son eines (tipus d'informe i pantalles de sistema): no porten
# segell. Qualsevol rajola NOVA en te automaticament, sense tocar cap llista.
$Script:AccionsSenseSegell = @('nou', 'seguiment', 'actextr', 'llicencia', 'mnstraspas', 'llicdb', 'config', 'editcataleg')

# Excepcio: dues eines ja escriuen la seva PROPIA marca quan han treballat de
# debo (la necessiten per anar en incremental), i aquella data es mes precisa que
# "has obert l'eina". Es llegeix primer i, si no hi es, es cau al registre.
$Script:SegellPropi = @{
    informesdb     = @{ Fitxer = 'informes-db.json';          Prop = 'actualitzat_el' }
    copiarinformes = @{ Fitxer = 'copia-informes-state.json'; Prop = 'copiat_el' }
}

function _EinesStatePath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'eines-state.json')
}

# Apunta que aquesta eina s'acaba de fer servir. No llanca mai: si la carpeta no
# hi es (unitat de xarxa fora de servei), el menu ha de seguir funcionant igual.
function _MarcaEinaUsada([string]$accio) {
    if ([string]::IsNullOrWhiteSpace($accio)) { return }
    try {
        $p = _EinesStatePath
        if ([string]::IsNullOrWhiteSpace($p)) { return }
        $dir = Split-Path -Parent $p
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $dades = [ordered]@{}
        $o = Read-JsonFile $p
        if ($null -ne $o) {
            foreach ($pr in $o.PSObject.Properties) { $dades[$pr.Name] = [string]$pr.Value }
        }
        $dades[$accio] = (Get-Date).ToString('o')
        # [pscustomobject] (i no el diccionari pelat): es l'idioma que ja fa
        # servir la resta del programa i a PowerShell 5.1 serialitza segur com un
        # objecte JSON, conservant l'ordre.
        Write-JsonFile $p ([pscustomobject]$dades) 5
    } catch { }
}

# El text del segell d'una eina.
function _LastRunEina([string]$accio) {
    if ([string]::IsNullOrWhiteSpace($accio)) { return '(mai)' }
    if ($Script:SegellPropi.Contains($accio) -and -not [string]::IsNullOrWhiteSpace($LocalActivitatsDir)) {
        $d = $Script:SegellPropi[$accio]
        $t = _LastRunText (Join-Path $LocalActivitatsDir $d.Fitxer) $d.Prop
        if ($t -ne '(mai)') { return $t }
    }
    return (_LastRunText (_EinesStatePath) $accio)
}

# Retorna @{ Action='nou'|'seguiment'|'actextr'; Cataleg=<FileInfo|$null> }.
# Per a 'nou', Cataleg es el .docx triat (ja no cal un segon pas de tria).
# Tancar la finestra (X) avorta (exit 0).
function Select-Mode {
    # Catalegs disponibles a ESTRUCTURALS (REQ1.json, TERMINI.json...). Es
    # descobreixen sols; els noms amics dels coneguts es defineixen mes avall.
    $catalegs = @(Get-Catalegs)
    $byName = @{}
    foreach ($c in $catalegs) { $byName[$c.BaseName] = $c }

    # Noms accentuats fets amb codepoint (Seguiment.ps1 no porta BOM: un literal
    # accentuat es corromp segons l'encoding amb que PowerShell 5.1 llegeix el
    # fitxer). U+00F3 = 'o' accent tancat; U+00E0 = 'a' accent obert.
    $aG = [char]0x00E0   # a accent obert
    $eG = [char]0x00E8   # e accent obert
    $oG = [char]0x00F3   # o accent tancat
    $oT = [char]0x00F3   # o accent tancat
    $ampliacio      = 'Ampliaci' + $oT + ' termini'
    $extraordinaria = 'Activitats extraordin' + $aG + 'ries'

    # Icones (emoji astral -> ConvertFromUtf32). Es pinten al xip granat suau.
    $icoNou = [System.Char]::ConvertFromUtf32(0x1F4DD)   # 📝
    $icoSeg = [System.Char]::ConvertFromUtf32(0x1F504)   # 🔄
    $icoTer = [System.Char]::ConvertFromUtf32(0x23F1)    # ⏱
    $icoExt = [System.Char]::ConvertFromUtf32(0x1F3AA)   # 🎪
    $icoLlic = [System.Char]::ConvertFromUtf32(0x1F4DC)  # rotlle: llicencia
    $icoMns  = [System.Char]::ConvertFromUtf32(0x1F504)  # fletxes: modificacio / traspas
    $mnsNom  = ('Modificaci' + $oG + ' No Substancial / Trasp' + $aG + 's')

    # Menu ORDENAT. Cada entrada: Action, Label (nom amic), Sub (descripcio
    # curta en gris), Icon (emoji del xip), Doc (xip del document a la dreta) i,
    # per a 'nou', el Cataleg (FileInfo). Els 'nou' nomes surten si el .docx hi es.
    $menu = New-Object System.Collections.ArrayList
    # L'ORDRE DEL MENU EL DECIDEIX L'USUARI i es aquest (agost 2026):
    #   1 Requeriment - Nou   2 Requeriment - Seguiment   3 Llicencia
    #   4 Activitats extraordinaries   5 MNS / Traspas   6 Ampliacio termini
    if ($byName.ContainsKey('REQ1'))    { [void]$menu.Add(@{ Action='nou'; Label='Requeriment - Nou'; Sub=('Cat' + $aG + 'leg de defici' + $eG + 'ncies'); Icon=$icoNou; Doc='REQ1'; Cataleg=$byName['REQ1'] }) }
    [void]$menu.Add(@{ Action='seguiment'; Label='Requeriment - Seguiment'; Sub='Sobre un informe ja fet'; Icon=$icoSeg; Doc=''; Cataleg=$null })
    # Llicencia: NO passa el cataleg (LLIC no es un cataleg de deficiencies sino
    # la capa propia de Llicencia sobre REQ1; vegeu Llicencia.ps1).
    $llicNom = 'Llic' + $eG + 'ncia (Annex II / LL Prov)'
    # 'Extra': un SEGON xip clicable a la mateixa fila, a l'esquerra del de
    # ✏️ LLIC. Obre la base de dades de llicencies (el que es recorda de cada
    # activitat per als informes seguents).
    [void]$menu.Add(@{ Action='llicencia'; Label=$llicNom; Sub='Requeriment i favorables'; Icon=$icoLlic; Doc='LLIC'; Cataleg=$null;
                       Extra=@{ Text='Dades'; Icon=([System.Char]::ConvertFromUtf32(0x1F5C2) + [char]0xFE0F); Action='llicdb' } })
    [void]$menu.Add(@{ Action='actextr'; Label=$extraordinaria; Sub='Decret 112/2010'; Icon=$icoExt; Doc='ACT_EXTR'; Cataleg=$null })
    # MNS / TRASPAS: ENTRADA PROPIA. Abans s'hi arribava des de dins de
    # Llicencia (el pas 1 oferia les cinc fases juntes) i no es veia des del
    # menu. Comparteixen capcalera, tramit i base de dades amb Llicencia -per
    # aixo passen pel mateix assistent-, pero son informes a part.
    [void]$menu.Add(@{ Action='mnstraspas'; Label=$mnsNom; Sub='Informes curts'; Icon=$icoMns; Doc='MNSTRAS'; Cataleg=$null })
    if ($byName.ContainsKey('TERMINI')) { [void]$menu.Add(@{ Action='nou'; Label=$ampliacio; Sub='Informe de cos fix'; Icon=$icoTer; Doc='TERMINI'; Cataleg=$byName['TERMINI'] }) }
    # Qualsevol altre cataleg no llistat (p.ex. un REQ2 nou) s'afegeix al final.
    foreach ($c in $catalegs) {
        if ($c.BaseName -in 'REQ1','TERMINI') { continue }
        [void]$menu.Add(@{ Action='nou'; Label=$c.BaseName; Sub=''; Icon=$icoNou; Doc=$c.BaseName; Cataleg=$c })
    }

    $form = _NewForm
    $form.Text = 'Informes Cornella - Pas 1'
    $form.StartPosition = 'CenterScreen'

    # Banda de capcalera GRANAT (color corporatiu + titol de l'app). Nomes
    # desplaca cap avall el punt de partida ($headerHeight): la resta del menu
    # (tots els botons, ja calculats amb $y +=) no s'ha de retocar.
    $headerHeight = 56
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Que vols fer?'
    $lbl.Location = New-Object System.Drawing.Point(20, (15 + $headerHeight))
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $fMain = New-Object System.Drawing.Font('Segoe UI', 12.5, [System.Drawing.FontStyle]::Bold)
    $fDet  = New-Object System.Drawing.Font('Segoe UI', 9.5,  [System.Drawing.FontStyle]::Regular)
    $fIcon = New-Object System.Drawing.Font('Segoe UI Emoji', 15, [System.Drawing.FontStyle]::Regular)
    $fEmoS = New-Object System.Drawing.Font('Segoe UI Emoji', 9, [System.Drawing.FontStyle]::Regular)
    $pencil = [string][char]0x270F + [char]0xFE0F   # emoji d'editar
    $flags = [System.Windows.Forms.TextFormatFlags]::NoPadding
    $flagsC = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::NoPadding
    $colGranat = [System.Drawing.Color]::FromArgb(166, 26, 47)
    $colSoft   = [System.Drawing.Color]::FromArgb(247, 231, 234)
    $colInk    = [System.Drawing.Color]::FromArgb(29, 39, 51)
    $colSub    = [System.Drawing.Color]::FromArgb(107, 116, 128)

    # Dibuix propietari del boto de generacio (protagonista): xip granat suau amb
    # icona a l'esquerra, titol + subtitol al centre-esquerra, i xip del document
    # a la dreta.
    $paintHandler = {
        param($sender, $e)
        $entry = $sender.Tag
        $g = $e.Graphics
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $rect = $sender.ClientRectangle
        $main = [string]$entry.Label
        $sub  = [string]$entry.Sub
        $ico  = [string]$entry.Icon
        $doc  = [string]$entry.Doc

        # Xip d'icona a l'esquerra.
        $chip = 42
        $cx = 12; $cy = [int](($rect.Height - $chip) / 2)
        $bSoft = New-Object System.Drawing.SolidBrush($colSoft)
        $g.FillRectangle($bSoft, $cx, $cy, $chip, $chip)
        $bSoft.Dispose()
        if ($ico) {
            $icoRect = New-Object System.Drawing.Rectangle($cx, ($cy - 1), $chip, $chip)
            [System.Windows.Forms.TextRenderer]::DrawText($g, $ico, $fIcon, $icoRect, $colGranat, $flagsC)
        }

        # ELS XIPS PRIMER, i despres el titol dins del que quedi: si es dibuixa
        # el titol sense limit, un nom llarg passa PER SOTA dels xips (el xip
        # "Dades" tapava el "LL Prov" de Llicencia). Amb EndEllipsis el text es
        # retalla amb punts suspensius i no pot solapar-se MAI, digui el que
        # digui i hi hagi els xips que hi hagi.
        # Xip del document a la dreta. Es CLICABLE: obre l'editor de catalegs
        # (hi dibuixem un emoji d'editar ✏️ i en guardem el rectangle per al
        # hit-test del clic, a $entry.DocChipRect).
        $entry.DocChipRect = $null
        $entry.ExtraChipRect = $null
        if (-not [string]::IsNullOrWhiteSpace($doc)) {
            $szP = [System.Windows.Forms.TextRenderer]::MeasureText($g, $pencil, $fEmoS, [System.Drawing.Size]::Empty, $flags)
            $szD = [System.Windows.Forms.TextRenderer]::MeasureText($g, $doc, $fDet, [System.Drawing.Size]::Empty, $flags)
            $pad = 9; $gap = 5
            $cw = $pad + $szP.Width + $gap + $szD.Width + $pad
            $chH = $szD.Height + 8
            $dx = $rect.Width - $cw - 14
            $dy = [int](($rect.Height - $chH) / 2)
            # En passar-hi el ratolí (ChipHover) es ressalta com un botó: fons més
            # intens + vora granat (el cursor passa a "mà" al MouseMove).
            $chipBg = if ($entry.ChipHover) { [System.Drawing.Color]::FromArgb(238, 208, 213) } else { $colSoft }
            $bD = New-Object System.Drawing.SolidBrush($chipBg)
            $g.FillRectangle($bD, $dx, $dy, $cw, $chH)
            $bD.Dispose()
            if ($entry.ChipHover) {
                $penH = New-Object System.Drawing.Pen($colGranat)
                $g.DrawRectangle($penH, $dx, $dy, ($cw - 1), ($chH - 1))
                $penH.Dispose()
            }
            [System.Windows.Forms.TextRenderer]::DrawText($g, $pencil, $fEmoS, (New-Object System.Drawing.Point(($dx + $pad), ($dy + 5))), $colGranat, $flags)
            [System.Windows.Forms.TextRenderer]::DrawText($g, $doc, $fDet, (New-Object System.Drawing.Point(($dx + $pad + $szP.Width + $gap), ($dy + 4))), $colGranat, $flags)
            $entry.DocChipRect = New-Object System.Drawing.Rectangle($dx, $dy, $cw, $chH)

            # Xip EXTRA (opcional), just a l'esquerra del del document. Mateix
            # aspecte i mateix hit-test; el seu rectangle va a $entry.ExtraChipRect.
            if ($null -ne $entry.Extra) {
                $et = [string]$entry.Extra.Text
                $ei = [string]$entry.Extra.Icon
                $szEI = [System.Windows.Forms.TextRenderer]::MeasureText($g, $ei, $fEmoS, [System.Drawing.Size]::Empty, $flags)
                $szET = [System.Windows.Forms.TextRenderer]::MeasureText($g, $et, $fDet, [System.Drawing.Size]::Empty, $flags)
                $ew = $pad + $szEI.Width + $gap + $szET.Width + $pad
                $ex = $dx - $ew - 8
                $exBg = if ($entry.ExtraHover) { [System.Drawing.Color]::FromArgb(238, 208, 213) } else { $colSoft }
                $bE = New-Object System.Drawing.SolidBrush($exBg)
                $g.FillRectangle($bE, $ex, $dy, $ew, $chH)
                $bE.Dispose()
                if ($entry.ExtraHover) {
                    $penE = New-Object System.Drawing.Pen($colGranat)
                    $g.DrawRectangle($penE, $ex, $dy, ($ew - 1), ($chH - 1))
                    $penE.Dispose()
                }
                [System.Windows.Forms.TextRenderer]::DrawText($g, $ei, $fEmoS, (New-Object System.Drawing.Point(($ex + $pad), ($dy + 5))), $colGranat, $flags)
                [System.Windows.Forms.TextRenderer]::DrawText($g, $et, $fDet, (New-Object System.Drawing.Point(($ex + $pad + $szEI.Width + $gap), ($dy + 4))), $colGranat, $flags)
                $entry.ExtraChipRect = New-Object System.Drawing.Rectangle($ex, $dy, $ew, $chH)
            }
        }

        # Titol + subtitol, ACOTATS pel xip mes a l'esquerra.
        $tx = $cx + $chip + 14
        $limit = $rect.Width - 14
        if ($null -ne $entry.ExtraChipRect) { $limit = $entry.ExtraChipRect.Left }
        elseif ($null -ne $entry.DocChipRect) { $limit = $entry.DocChipRect.Left }
        $ampleText = [Math]::Max(40, $limit - $tx - 10)
        $flagsT = $flags -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis
        if (-not [string]::IsNullOrWhiteSpace($sub)) {
            $rT = New-Object System.Drawing.Rectangle($tx, 11, $ampleText, 24)
            $rS = New-Object System.Drawing.Rectangle($tx, 35, $ampleText, 20)
            [System.Windows.Forms.TextRenderer]::DrawText($g, $main, $fMain, $rT, $colInk, $flagsT)
            [System.Windows.Forms.TextRenderer]::DrawText($g, $sub,  $fDet,  $rS, $colSub, $flagsT)
        } else {
            $szM = [System.Windows.Forms.TextRenderer]::MeasureText($g, $main, $fMain, [System.Drawing.Size]::Empty, $flags)
            $rT = New-Object System.Drawing.Rectangle($tx, [int](($rect.Height - $szM.Height) / 2), $ampleText, $szM.Height)
            [System.Windows.Forms.TextRenderer]::DrawText($g, $main, $fMain, $rT, $colInk, $flagsT)
        }
    }

    $result = @{ Choice = $null }

    # ELS DOS DOCUMENTS QUE SON DE TOTS ELS INFORMES: la capcalera i les
    # conclusions. No son de cap tipus d'informe en concret -per aixo no tenen
    # xip a cap rajola- i fins ara no s'hi podia arribar des d'enlloc. Van al
    # costat de "Que vols fer?", com va demanar l'usuari.
    #
    # VA DESPRES DE $result I AMB .GetNewClosure(): un scriptblock SENSE closure
    # no veu els LOCALS de la funcio que el crea (nomes l'ambit de l'script), o
    # sigui que "$result.Choice = ..." queia sobre $null, el menu es tancava i el
    # programa sortia sense fer res. Va passar de debo.
    #
    # I SENSE EMOJI: un LinkLabel te UNA sola lletra per a tot el text, i amb la
    # Segoe UI del programa el llapis surt com un quadrat. Als xips de les
    # rajoles si que hi es perque alla el dibuixem a part, amb Segoe UI Emoji.
    $xComuns = 130
    foreach ($d in @(
        @{ Doc = '0 CAPCALERA';   Text = ('Cap' + [char]0x00E7 + 'alera') },
        @{ Doc = '0 CONCLUSIONS'; Text = 'Conclusions' })) {
        $ll = New-Object System.Windows.Forms.LinkLabel
        $ll.Text = [string]$d.Text
        $ll.Location = New-Object System.Drawing.Point($xComuns, (15 + $headerHeight))
        $ll.AutoSize = $true
        $ll.Font = $fDet
        $ll.LinkColor = $colGranat
        $ll.ActiveLinkColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
        $ll.LinkBehavior = 'HoverUnderline'
        $ll.Tag = [string]$d.Doc
        [void]$form.Controls.Add($ll)
        $ll.add_LinkClicked({
            param($snd, $ev)
            $result.Choice = @{ Action = 'editcataleg'; Doc = [string]$snd.Tag; Cataleg = $null }
            $form.DialogResult = 'OK'
            $form.Close()
        }.GetNewClosure())
        $xComuns += $ll.PreferredWidth + 18
    }
    $y = 45 + $headerHeight
    foreach ($entry in $menu) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = ''
        $btn.Tag = $entry
        $btn.Location = New-Object System.Drawing.Point(20, $y)
        $btn.Size = New-Object System.Drawing.Size(560, 62)
        $btn.FlatStyle = 'Flat'
        $btn.BackColor = [System.Drawing.Color]::White
        $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 219, 225)
        $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(250, 240, 242)
        $btn.add_Paint($paintHandler)
        # Clic amb coordenades: si es damunt del xip del document (✏️), obre
        # l'editor de catalegs; si no, tria el tipus d'informe com sempre.
        $btn.add_MouseClick({
            param($s, $e)
            $en = $s.Tag
            $rc = $en.DocChipRect
            $rx = $en.ExtraChipRect
            if ($null -ne $rx -and $rx.Contains($e.Location)) {
                $result.Choice = @{ Action = [string]$en.Extra.Action; Doc = [string]$en.Doc; Cataleg = $null }
            } elseif ($null -ne $rc -and $rc.Contains($e.Location)) {
                $result.Choice = @{ Action = 'editcataleg'; Doc = [string]$en.Doc; Cataleg = $null }
            } else {
                $result.Choice = $en
            }
            $form.DialogResult = 'OK'
            $form.Close()
        }.GetNewClosure())
        # Feedback de que el xip ✏️ es clicable: cursor "mà" i ressaltat quan el
        # ratolí hi és a sobre (nomes es repinta quan l'estat de hover canvia).
        $btn.add_MouseMove({
            param($s, $e)
            $en = $s.Tag
            $rc = $en.DocChipRect
            $rx = $en.ExtraChipRect
            $over  = ($null -ne $rc -and $rc.Contains($e.Location))
            $overX = ($null -ne $rx -and $rx.Contains($e.Location))
            if ($over -ne [bool]$en.ChipHover -or $overX -ne [bool]$en.ExtraHover) {
                $en.ChipHover = $over
                $en.ExtraHover = $overX
                $s.Cursor = if ($over -or $overX) { [System.Windows.Forms.Cursors]::Hand } else { [System.Windows.Forms.Cursors]::Default }
                $s.Invalidate()
            }
        }.GetNewClosure())
        $btn.add_MouseLeave({
            param($s, $e)
            $en = $s.Tag
            if ([bool]$en.ChipHover -or [bool]$en.ExtraHover) {
                $en.ChipHover = $false; $en.ExtraHover = $false
                $s.Cursor = [System.Windows.Forms.Cursors]::Default; $s.Invalidate()
            }
        }.GetNewClosure())
        [void]$form.Controls.Add($btn)
        $y += 70
    }

    # ---- Eines (separades dels tipus d'informe) ----------------------------
    $y += 6
    $sepEines = New-Object System.Windows.Forms.Label
    $sepEines.Text = 'EINES'
    $sepEines.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $sepEines.ForeColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $sepEines.Location = New-Object System.Drawing.Point(20, $y)
    $sepEines.AutoSize = $true
    [void]$form.Controls.Add($sepEines)
    $y += 24

    # EINES: rajoles compactes en una fila (emoji a dalt + etiqueta petita a
    # sota), segons el disseny. Comportament per rajola: 'action' tanca el menu
    # amb l'accio; 'url' obre l'enllac public SENSE tancar el menu (precintades).
    # (Les eines de la base d'informes van al seu propi apartat INFORMES, sota.)
    $urlPrec = 'https://xexifm.github.io/informes-Cornella/precintades.html'
    $tiPin   = [System.Char]::ConvertFromUtf32(0x1F4CD)   # 📍
    $tiLock  = [System.Char]::ConvertFromUtf32(0x1F512)   # 🔒
    $tiBox   = [System.Char]::ConvertFromUtf32(0x1F5C3)   # 🗃
    $tiClip  = [System.Char]::ConvertFromUtf32(0x1F4CB)   # 📋
    $tiInbox = [System.Char]::ConvertFromUtf32(0x1F4E5)   # 📥
    $tiCopy  = [System.Char]::ConvertFromUtf32(0x1F4C1)   # 📁
    $tiCheck = [System.Char]::ConvertFromUtf32(0x2705)    # ✅
    $tiCal   = [System.Char]::ConvertFromUtf32(0x1F4C5)   # 📅
    $tiPdf   = [System.Char]::ConvertFromUtf32(0x1F4C4)   # 📄
    $tiMail  = [System.Char]::ConvertFromUtf32(0x1F4E7)   # 📧
    $tiSend  = [System.Char]::ConvertFromUtf32(0x1F4E4)   # 📤
    $tiList  = [System.Char]::ConvertFromUtf32(0x1F4CA)   # 📊
    $tiMap   = [System.Char]::ConvertFromUtf32(0x1F5FA)   # 🗺
    $tiBell  = [System.Char]::ConvertFromUtf32(0x1F514)   # 🔔
    # EINES: utilitats generals.
    $tools = @(
        @{ Emoji = $tiPin;   Label = 'Generar ruta';           Kind = 'action'; Action = 'ruta' }
        @{ Emoji = $tiMap;   Label = 'Coordenades';            Kind = 'action'; Action = 'coordenades' }
        # 'Action' tambe a la rajola d'enllac: no despatxa res, pero es la clau
        # del seu segell d'ultima execucio.
        @{ Emoji = $tiLock;  Label = 'Activitats precintades'; Kind = 'url';    Action = 'precintades'; Url = $urlPrec }
        @{ Emoji = $tiCal;   Label = ('Controls peri' + [char]0x00F2 + 'dics'); Kind = 'action'; Action = 'controlsperiodics' }
        @{ Emoji = $tiBell;  Label = 'Recordatoris'; Kind = 'action'; Action = 'recordatoris' }
    )
    # INFORMES: eines de la base d'informes + conversio a PDF.
    $reports = @(
        @{ Emoji = $tiBox;   Label = 'Actualitzar base'; Kind = 'action'; Action = 'informesdb' }
        @{ Emoji = $tiClip;  Label = 'Editar base';      Kind = 'action'; Action = 'informesdbedit' }
        @{ Emoji = $tiCopy;  Label = 'Copiar informes';  Kind = 'action'; Action = 'copiarinformes' }
        @{ Emoji = $tiPdf;   Label = 'Word a PDF';       Kind = 'action'; Action = 'convertirpdf' }
    )
    # GIA: eines que parlen de la base de dades d'ACTIVITATS (el GIA), no dels
    # informes. 'Comprovar Excel' era a INFORMES pero el seu tema es el GIA.
    $gia = @(
        @{ Emoji = $tiCheck; Label = 'Comprovar Excel'; Kind = 'action'; Action = 'comprovarexcel' }
        @{ Emoji = $tiList;  Label = 'Seguiment';       Kind = 'action'; Action = 'seguimentgia' }
    )
    # MOBIL: eines de l'app del mobil.
    $mobil = @(
        @{ Emoji = $tiMail;  Label = 'Textos del correu'; Kind = 'action'; Action = 'emailtextos' }
        @{ Emoji = $tiSend;  Label = 'Enviar correu';     Kind = 'action'; Action = 'enviarcorreu' }
        @{ Emoji = $tiInbox; Label = ('Revisar m' + [char]0x00F2 + 'bil'); Kind = 'action'; Action = 'revisarmobil' }
    )
    $fTileIco   = New-Object System.Drawing.Font('Segoe UI Emoji', 14, [System.Drawing.FontStyle]::Regular)
    $fTileTxt   = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular)
    $tileBorder = [System.Drawing.Color]::FromArgb(214, 219, 225)
    $tileTxtCol = [System.Drawing.Color]::FromArgb(107, 116, 128)
    $tilePaint = {
        param($s, $e)
        $t = $s.Tag
        $g = $e.Graphics
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $rc = $s.ClientRectangle
        $flC = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::NoPadding
        $flW = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::WordBreak -bor [System.Windows.Forms.TextFormatFlags]::NoPadding
        $emRect = New-Object System.Drawing.Rectangle(0, 6, $rc.Width, 22)
        [System.Windows.Forms.TextRenderer]::DrawText($g, $t.Emoji, $fTileIco, $emRect, [System.Drawing.Color]::Black, $flC)
        $lbRect = New-Object System.Drawing.Rectangle(2, 29, ($rc.Width - 4), ($rc.Height - 31))
        [System.Windows.Forms.TextRenderer]::DrawText($g, $t.Label, $fTileTxt, $lbRect, $tileTxtCol, $flW)
    }.GetNewClosure()
    $tileClick = {
        param($s, $e)
        $t = $s.Tag
        if ($t.Kind -eq 'url') {
            try {
                Start-Process $t.Url | Out-Null
                # Aquesta rajola NO tanca el menu, o sigui que no passa pel
                # despatxador: el segell s'apunta i es refresca aqui mateix.
                _MarcaEinaUsada ([string]$t.Action)
                if ($null -ne $t.StampLabel) { $t.StampLabel.Text = [string](_LastRunEina ([string]$t.Action)) }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("No s'ha pogut obrir l'enllac:`n$($t.Url)", 'Eina', 'OK', 'Error') | Out-Null
            }
        } else {
            $result.Choice = @{ Action = $t.Action; Cataleg = $null }
            $form.DialogResult = 'OK'
            $form.Close()
        }
    }.GetNewClosure()
    $tileW = 80; $tileH = 58; $tileGap = 7
    # Sota CADA rajola, en petit, l'ultima vegada que s'ha fet servir l'eina
    # ('(mai)' si encara no). El segell es llegeix per l'ACCIO de la rajola (la
    # clau del registre), no per la posicio dins de la fila: abans els indexs
    # anaven a pinyo fix contra una fila concreta i, en moure 'Comprovar Excel'
    # de fila, el segell hauria anat a la rajola equivocada.
    $fStamp = New-Object System.Drawing.Font('Segoe UI', 7)
    $colStamp = [System.Drawing.Color]::FromArgb(120, 128, 138)
    # Dibuixa una fila de rajoles amb el seu segell a l'alcada $y actual i retorna
    # la $y seguent (helper unic: el fan servir les quatre files).
    $addTileRow = {
        param($items, $yRow)
        $tx = 20
        foreach ($tool in $items) {
            $tb = New-Object System.Windows.Forms.Button
            $tb.Text = ''
            $tb.Tag = $tool
            $tb.Location = New-Object System.Drawing.Point($tx, $yRow)
            $tb.Size = New-Object System.Drawing.Size($tileW, $tileH)
            $tb.FlatStyle = 'Flat'
            $tb.BackColor = [System.Drawing.Color]::White
            $tb.FlatAppearance.BorderColor = $tileBorder
            $tb.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(250, 240, 242)
            $tb.add_Paint($tilePaint)
            $tb.add_Click($tileClick)
            [void]$form.Controls.Add($tb)

            $lblS = New-Object System.Windows.Forms.Label
            $lblS.Text = [string](_LastRunEina ([string]$tool.Action))
            $lblS.Font = $fStamp
            $lblS.ForeColor = $colStamp
            $lblS.TextAlign = 'MiddleCenter'
            $lblS.Location = New-Object System.Drawing.Point($tx, ($yRow + $tileH + 2))
            $lblS.Size = New-Object System.Drawing.Size($tileW, 14)
            [void]$form.Controls.Add($lblS)
            # El guardem a la propia rajola: la d'enllac (precintades) no tanca el
            # menu i s'ha de poder refrescar el seu segell alli mateix.
            $tool.StampLabel = $lblS

            $tx += $tileW + $tileGap
        }
        return ($yRow + $tileH + 20)
    }.GetNewClosure()
    $y = & $addTileRow $tools $y

    # ---- INFORMES (base d'informes) ----------------------------------------
    $sepInformes = New-Object System.Windows.Forms.Label
    $sepInformes.Text = 'INFORMES'
    $sepInformes.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $sepInformes.ForeColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $sepInformes.Location = New-Object System.Drawing.Point(20, $y)
    $sepInformes.AutoSize = $true
    [void]$form.Controls.Add($sepInformes)
    $y += 24

    $y = & $addTileRow $reports $y

    # ---- GIA (base de dades d'activitats) ----------------------------------
    $sepGia = New-Object System.Windows.Forms.Label
    $sepGia.Text = 'GIA'
    $sepGia.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $sepGia.ForeColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $sepGia.Location = New-Object System.Drawing.Point(20, $y)
    $sepGia.AutoSize = $true
    [void]$form.Controls.Add($sepGia)
    $y += 24
    $y = & $addTileRow $gia $y

    # ---- MOBIL (app del mobil) ---------------------------------------------
    $sepMobil = New-Object System.Windows.Forms.Label
    $sepMobil.Text = 'M' + [char]0x00D2 + 'BIL'
    $sepMobil.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $sepMobil.ForeColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $sepMobil.Location = New-Object System.Drawing.Point(20, $y)
    $sepMobil.AutoSize = $true
    [void]$form.Controls.Add($sepMobil)
    $y += 24
    $y = & $addTileRow $mobil $y
    $y += 4

    # (Configuracio i Ajuda ja no son botons grans: van DISCRETS a la cantonada
    #  de la banda granat, mes avall.)
    $urlAjuda = 'https://github.com/xexifm/informes-cornella/blob/main/LLEGEIX-ME.md'

    $form.ClientSize = New-Object System.Drawing.Size(600, ($y + 12))

    # Banda de capcalera GRANAT amb escut blanc (helper comu del redisseny).
    # S'afegeix al final (Dock=Top) per no desplacar els controls ja posicionats.
    $subTitle = 'Ajuntament de Cornell' + [char]0x00E0 + ' de Llobregat'
    $band = _AddBrandHeader $form "Generador d'informes" $subTitle $headerHeight

    # Botons DISCRETS a la cantonada dreta de la banda: Ajuda (?) i Configuracio
    # (rosca). Fons granat una mica mes clar, text blanc, sense vora. Ancorats a
    # la dreta perque segueixin la cantonada si es maximitza.
    $wForm = $form.ClientSize.Width
    $fBandIco = New-Object System.Drawing.Font('Segoe UI Emoji', 11, [System.Drawing.FontStyle]::Regular)
    $btnAjuda = New-Object System.Windows.Forms.Button
    $btnAjuda.Text = [string][char]0x2753
    $btnAjuda.Font = $fBandIco
    $btnAjuda.Size = New-Object System.Drawing.Size(30, 30)
    $btnAjuda.Location = New-Object System.Drawing.Point(($wForm - 42), 13)
    $btnAjuda.Anchor = 'Top,Right'
    $btnAjuda.FlatStyle = 'Flat'
    $btnAjuda.ForeColor = [System.Drawing.Color]::White
    $btnAjuda.BackColor = [System.Drawing.Color]::FromArgb(150, 45, 60)
    $btnAjuda.FlatAppearance.BorderSize = 0
    $btnAjuda.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $btnAjuda.add_Click({
        try { Start-Process $urlAjuda | Out-Null } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut obrir l'enllac:`n$urlAjuda", 'Ajuda', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    [void]$band.Controls.Add($btnAjuda)

    # CARPETA DELS INFORMES GENERATS. La ruta surt de _ResolveOutputDir, o sigui
    # que es EXACTAMENT la que hi ha a Configuracio (i el respatller local si
    # aquella no s'hi pot arribar): aqui no hi ha cap ruta escrita.
    #
    # El menu NO es tanca: obrir una carpeta no es triar cap opcio.
    #
    # L'EMOJI DE CARPETA ES ASTRAL (U+1F4C1): [char] es de 16 bits i no hi cap
    # -aixo ja va deixar el programa sense arrencar un cop-, per aixo va amb
    # ConvertFromUtf32. Ho vigila una prova.
    $btnCarpeta = New-Object System.Windows.Forms.Button
    $btnCarpeta.Text = [System.Char]::ConvertFromUtf32(0x1F4C1)
    $btnCarpeta.Font = $fBandIco
    $btnCarpeta.Size = New-Object System.Drawing.Size(30, 30)
    $btnCarpeta.Location = New-Object System.Drawing.Point(($wForm - 114), 13)
    $btnCarpeta.Anchor = 'Top,Right'
    $btnCarpeta.FlatStyle = 'Flat'
    $btnCarpeta.ForeColor = [System.Drawing.Color]::White
    $btnCarpeta.BackColor = [System.Drawing.Color]::FromArgb(150, 45, 60)
    $btnCarpeta.FlatAppearance.BorderSize = 0
    $btnCarpeta.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $btnCarpeta.add_Click({
        try {
            $carpeta = [string](_ResolveOutputDir)
            if ([string]::IsNullOrWhiteSpace($carpeta)) { throw "no hi ha cap carpeta de sortida configurada" }
            if (-not (Test-Path -LiteralPath $carpeta)) { throw ("no existeix: " + $carpeta) }
            Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $carpeta + '"') | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                ("No s'ha pogut obrir la carpeta dels informes:`n`n" + $_.Exception.Message +
                 "`n`nLa pots canviar al boto de Configuracio."),
                'Informes generats', 'OK', 'Warning') | Out-Null
        }
    }.GetNewClosure())
    [void]$band.Controls.Add($btnCarpeta)
    $ttBand = New-Object System.Windows.Forms.ToolTip
    $ttBand.SetToolTip($btnCarpeta, 'Obre la carpeta dels informes generats')
    $ttBand.SetToolTip($btnAjuda, 'Ajuda')

    $btnConfig = New-Object System.Windows.Forms.Button
    $btnConfig.Text = [string][char]0x2699
    $btnConfig.Font = $fBandIco
    $btnConfig.Size = New-Object System.Drawing.Size(30, 30)
    $btnConfig.Location = New-Object System.Drawing.Point(($wForm - 78), 13)
    $btnConfig.Anchor = 'Top,Right'
    $btnConfig.FlatStyle = 'Flat'
    $btnConfig.ForeColor = [System.Drawing.Color]::White
    $btnConfig.BackColor = [System.Drawing.Color]::FromArgb(150, 45, 60)
    $btnConfig.FlatAppearance.BorderSize = 0
    $btnConfig.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $btnConfig.add_Click({
        $result.Choice = @{ Action = 'config'; Cataleg = $null }
        $form.DialogResult = 'OK'
        $form.Close()
    }.GetNewClosure())
    [void]$band.Controls.Add($btnConfig)
    $ttBand.SetToolTip($btnConfig, 'Configuracio')

    $res = $form.ShowDialog()
    if ($res -ne 'OK' -or $null -eq $result.Choice) { exit 0 }
    $ch = $result.Choice
    return @{ Action = $ch.Action; Cataleg = $ch.Cataleg; Doc = $ch.Doc }
}
