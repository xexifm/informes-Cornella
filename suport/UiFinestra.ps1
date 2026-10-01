#requires -Version 5.1
<#
.SYNOPSIS
  Que una finestra HI CAPIGA sempre: scroll vertical i ajust a la pantalla.

.DESCRIPTION
  Problema real: en una pantalla mes baixa (el PC de casa, o el Windows amb
  escalat al 125%), diverses pantalles del programa son MES ALTES que l'area de
  treball. Llavors la part de baix -que es on van els botons Enrere/Seguent- no
  es pot veure ni arribar-hi: la finestra no es pot encongir perque te
  MinimumSize, i el que sobresurt queda fora de la pantalla.

  Aixo es arregla en DOS temps, i calen tots dos:

    1. AutoScroll a TOTES les finestres: si s'encongeixen i algun control queda
       per sota, surt la barra VERTICAL i s'hi arriba. Es deixa que el WinForms
       calculi sol la zona a recorrer, aixi una graella Dock='Fill' segueix
       encongint-se com sempre en lloc d'estrenar una barra que no calia.

    2. Nomes quan la finestra NO hi cap: encongir-la fins a l'area i, ABANS,
       BAIXAR-NE EL MinimumSize -si no, el Windows es nega a encongir-la-. En
       aquest cas s'hi fixa AutoScrollMinSize = l'alcada de DISSENY, que es
       l'unica manera de garantir que s'arriba a tot el que hi havia, tambe al
       que estigui ancorat a baix (que si no puja i es comprimeix). L'amplada es
       deixa a 0: els controls ancorats a la dreta ja s'estrenyen sols i posar-hi
       l'amplada de disseny trauria una barra HORITZONTAL que no cal.

  NO ES CARREGA RES EN AQUEST FITXER: nomes defineix funcions. Es a posta,
  perque el fan servir DOS PROCESSOS -el programa (via UiComuns.ps1) i el
  planificador de rutes (rutes/Ruta.ps1, que no carrega UiComuns perque aquell
  si que te efectes en carregar-se: AppUserModelID, icona...)-. Un modul amb
  efectes no es pot compartir entre processos sense arrossegar-los.
#>

# ----------------------------------------------------------------------------
# La DECISIO es pura: rebre mides i tornar què s'ha d'aplicar. Es prova a Linux.
# ----------------------------------------------------------------------------
# Entrada: la mida que vol la finestra, el seu MinimumSize i l'area de treball
# de la pantalla (la de debo, sense la barra de tasques).
# Sortida: @{ W; H; MinW; MinH; X; Y; Cal } -Cal = $false si ja hi cabia i no
# s'ha de tocar res-.
#
# Regles:
#   - el MinimumSize no pot passar de l'area: si hi passa, el Windows no deixa
#     encongir la finestra i tot plegat no serveix;
#   - la mida es retalla a l'area;
#   - i la posicio es corre perque la finestra quedi SENCERA a dins (una finestra
#     centrada que sobresurt per baix tambe sobresurt per dalt, i llavors ni la
#     barra de titol es pot agafar).
function _MidaFinestraDinsPantalla([int]$w, [int]$h, [int]$minW, [int]$minH,
                                   [int]$x, [int]$y,
                                   [int]$areaX, [int]$areaY, [int]$areaW, [int]$areaH) {
    $nMinW = if ($minW -gt $areaW) { $areaW } else { $minW }
    $nMinH = if ($minH -gt $areaH) { $areaH } else { $minH }
    $nW = if ($w -gt $areaW) { $areaW } else { $w }
    $nH = if ($h -gt $areaH) { $areaH } else { $h }
    if ($nW -lt $nMinW) { $nW = $nMinW }
    if ($nH -lt $nMinH) { $nH = $nMinH }

    $nX = $x
    $nY = $y
    if (($nX + $nW) -gt ($areaX + $areaW)) { $nX = $areaX + $areaW - $nW }
    if (($nY + $nH) -gt ($areaY + $areaH)) { $nY = $areaY + $areaH - $nH }
    if ($nX -lt $areaX) { $nX = $areaX }
    if ($nY -lt $areaY) { $nY = $areaY }

    return @{
        W = $nW; H = $nH; MinW = $nMinW; MinH = $nMinH; X = $nX; Y = $nY
        Cal = (($nW -ne $w) -or ($nH -ne $h) -or ($nMinW -ne $minW) -or
               ($nMinH -ne $minH) -or ($nX -ne $x) -or ($nY -ne $y))
    }
}

# ----------------------------------------------------------------------------
# L'APLICACIO (WinForms). Es crida des del Shown, quan la disposicio ja es
# definitiva: abans, el ClientSize encara pot canviar.
# ----------------------------------------------------------------------------
# L'alcada de DISSENY es la que te la finestra en obrir-se, ABANS de retallar-la:
# es la que ha de poder recorrer la barra de desplacament.
function _AjustaFinestraAPantalla($f) {
    if ($null -eq $f) { return }
    try {
        # L'alcada de DISSENY s'ha de llegir ABANS de tocar res.
        $altDisseny = [int]$f.ClientSize.Height

        # Scroll SEMPRE: si l'usuari encongeix la finestra i algun control queda
        # per sota, surt la barra i s'hi pot arribar. Amb AutoScrollMinSize a
        # zero, el WinForms calcula la zona a recorrer dels controls mateixos, o
        # sigui que una graella Dock='Fill' segueix ENCONGINT-SE com fins ara i
        # no apareix cap barra que abans no hi era.
        $f.AutoScroll = $true

        $area = ([System.Windows.Forms.Screen]::FromControl($f)).WorkingArea
        $r = _MidaFinestraDinsPantalla ([int]$f.Width) ([int]$f.Height) `
                                       ([int]$f.MinimumSize.Width) ([int]$f.MinimumSize.Height) `
                                       ([int]$f.Left) ([int]$f.Top) `
                                       ([int]$area.X) ([int]$area.Y) ([int]$area.Width) ([int]$area.Height)
        if (-not $r.Cal) { return }

        # Aqui si: la finestra NO hi cabia i la retallem. Llavors s'hi fixa
        # l'alcada de disseny com a zona recorrible, que es l'unica manera de
        # garantir que s'arriba a TOT el que hi havia -tambe al que estigui
        # ancorat a baix, que si no simplement pujaria i es comprimiria-.
        $f.AutoScrollMinSize = New-Object System.Drawing.Size(0, $altDisseny)
        # El MinimumSize PRIMER: si no, el Windows no deixa encongir la finestra.
        $f.MinimumSize = New-Object System.Drawing.Size([int]$r.MinW, [int]$r.MinH)
        $f.Size = New-Object System.Drawing.Size([int]$r.W, [int]$r.H)
        $f.Location = New-Object System.Drawing.Point([int]$r.X, [int]$r.Y)
    } catch { }
}

# Estil de boto PRIMARI (granat ple, text blanc) i SECUNDARI (blanc, text/vora
# granat). Reutilitzables a totes les pantalles del redisseny.
#
# Viuen AQUI i no a UiComuns.ps1 perque els fa servir _AddPeuBotons, que tambe
# corre al proces de rutes. El granat ($Script:BrandMaroon) el defineix
# UiComuns en carregar-se; al proces de rutes no hi es i _AddPeuBotons deixa els
# botons amb l'aspecte del sistema, com hi eren.
function _StylePrimaryButton($btn) {
    $btn.FlatStyle = 'Flat'
    $btn.BackColor = $Script:BrandMaroon
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatAppearance.BorderSize = 0
    $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $btn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $btn.Cursor = 'Hand'
}
# Boto d'ACCENT amb un color propi (blau mari per confirmar, vermell per
# descartar...). Mateixa carcassa que _StylePrimaryButton: aixi el color es
# l'unica cosa que canvia i no hi ha una tercera copia de l'estil escampada.
function _StyleAccentButton($btn, $fons, $fonsHover) {
    $btn.FlatStyle = 'Flat'
    $btn.BackColor = $fons
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatAppearance.BorderSize = 0
    if ($null -ne $fonsHover) { $btn.FlatAppearance.MouseOverBackColor = $fonsHover }
    $btn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $btn.Cursor = 'Hand'
}
function _StyleSecondaryButton($btn) {
    $btn.FlatStyle = 'Flat'
    $btn.BackColor = [System.Drawing.Color]::White
    $btn.ForeColor = $Script:BrandMaroon
    $btn.FlatAppearance.BorderColor = $Script:BrandMaroon
    $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(247, 231, 234)
    $btn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Regular)
    $btn.Cursor = 'Hand'
}

# ----------------------------------------------------------------------------
# EL PEU DE BOTONS, igual a totes les finestres
# ----------------------------------------------------------------------------
# Abans cada finestra se'l feia a ma: 18 "Enrere", 11 "Continuar/Seguent" i 21
# "Cancel.lar/Tancar", cadascun amb la seva mida (28, 30, 32 o 34 d'alt), la
# seva posicio, el seu estil (o cap: n'hi havia sense estil) i el sortir a
# l'esquerra en unes finestres i a la dreta en d'altres.
#
# LA CONVENCIO, una per a totes:
#   - a l'ESQUERRA, el que fa SORTIR o TORNAR (Enrere, Tancar, Cancel.lar) i, al
#     costat, les accions auxiliars (Marcar-ho tot, Exportar, Esborrar...);
#   - a la DRETA, el que fa AVANCAR; l'accio principal, la del tot a la dreta i
#     en granat ple;
#   - 32 d'alt, 15 de marge, 10 entre botons i l'amplada que demana el text;
#   - Intro = l'accio principal, Esc = sortir (qui ho vol, ho diu a l'spec).
#
# Cada boto es un hashtable:
#   Nom        clau amb que es torna (per activar-lo, desactivar-lo...)
#   Text       el que hi posa
#   Estil      'primari' | 'secundari' (per defecte) | 'accent'
#   Fons, FonsHover  nomes per a 'accent'
#   Resultat   DialogResult ('OK', 'Cancel', 'Retry'...), si en porta
#   Clic       scriptblock del Click, si en porta
#   Intro/Esc  $true -> AcceptButton / CancelButton de la finestra
#   Ample      amplada fixa (si no, la del text)
#   Icona      un emoji o simbol al davant del text (vegeu _PosaIcona)
# Torna un hashtable Nom -> boto.
#
# $form es la finestra (per a l'Intro i l'Esc; en una pestanya, la pestanya);
# $pare, on van els botons (la finestra o un panell de peu); $y, la fila.
# -Ancorat els enganxa a baix (per a finestres que es poden fer mes grans).

# EMOJIS I SIMBOLS DINS D'UN BOTO O D'UNA ETIQUETA (setembre 2026).
# Un control de WinForms te UNA sola lletra per a tot el text, i la del programa
# (Segoe UI) no porta els emojis ni una colla de simbols: "Obre la norma" amb la
# cadena d'enllac, "Recuperar dades" amb la fletxa en cercle, els vistos i els
# avisos de Configuracio... sortien amb un QUADRAT. Els que si que es veien son
# els que es dibuixen amb 'Segoe UI Emoji': les icones de les rajoles del menu i
# els botons d'un sol emoji de la banda granat.
#
# _PosaIcona fa el mateix per a qualsevol control: dibuixa el simbol a part,
# amb aquella lletra i el color del text, en una imatge que es posa AL COSTAT
# del text. Les fletxes <- -> i les de pujar i baixar si que hi son a la Segoe
# UI i no cal (les fan servir tots els peus de botons). Hi ha guard
# (06-guards.ps1): cap simbol mes enlla d'aquelles fletxes pot anar al .Text
# d'un control si no passa per aqui o per la lletra d'emojis.
$Script:IconaLletra = 'Segoe UI Emoji'

# EMOJIS EN COLOR (octubre 2026). El GDI de WinForms -TextRenderer i
# DrawString- NO sap pintar les lletres de colors: amb 'Segoe UI Emoji' tots
# els emojis del programa sortien com un contorn d'un sol color, i l'usuari
# veia en un esbos fet amb el navegador els mateixos emojis en color ("jo no
# els veig aixi"). Per aixo cada emoji que fa servir la interficie es una
# imatge PNG a suport\emojis\ (64x64, fons transparent, generades amb Noto
# Color Emoji; vegeu el LLEGEIX-ME d'alla) i es dibuixa la imatge.
#
# El nom del fitxer surt de l'emoji (_EmojiFitxer, pura): els punts de codi
# en hexadecimal, sense el selector de variant FE0F. Si un emoji no te imatge
# es dibuixa amb la lletra com sempre: un emoji nou mai no queda en blanc. Hi
# ha guard (06-guards.ps1) que cada emoji del menu tingui la seva imatge.
#
# Els SIMBOLS (vist, avis, la i d'informacio, fletxes) NO hi son a posta: van
# del color del text i en un boto han de seguir el color del boto.
$Script:EmojiDir = Join-Path $PSScriptRoot 'emojis'
$Script:EmojiCache = @{}

function _EmojiFitxer([string]$emoji) {
    if ([string]::IsNullOrEmpty($emoji)) { return '' }
    $cps = New-Object System.Collections.ArrayList
    try {
        for ($i = 0; $i -lt $emoji.Length; $i++) {
            $cp = [char]::ConvertToUtf32($emoji, $i)
            if ([char]::IsHighSurrogate($emoji[$i])) { $i++ }
            if ($cp -eq 0xFE0F) { continue }
            [void]$cps.Add(('{0:x}' -f $cp))
        }
    } catch { return '' }
    if ($cps.Count -eq 0) { return '' }
    return (($cps -join '-') + '.png')
}

# La imatge de l'emoji, o $null si no n'hi ha. Es llegeix UN cop (cache).
# A MEMORIA i no amb Image.FromFile: aquell deixa el fitxer AGAFAT mentre el
# programa es obert, i el clone viu en una unitat de xarxa on Actualitzar.bat
# l'ha de poder sobreescriure.
function _EmojiImatge([string]$emoji) {
    $nom = _EmojiFitxer $emoji
    if ([string]::IsNullOrEmpty($nom)) { return $null }
    if ($Script:EmojiCache.ContainsKey($nom)) { return $Script:EmojiCache[$nom] }
    $img = $null
    $ruta = Join-Path $Script:EmojiDir $nom
    if (Test-Path -LiteralPath $ruta) {
        $ms = $null; $tmp = $null
        try {
            $ms = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($ruta))
            $tmp = [System.Drawing.Image]::FromStream($ms)
            # Copia propia: la del FromStream necessita el stream obert per sempre.
            $img = New-Object System.Drawing.Bitmap($tmp)
        } catch { $img = $null } finally {
            if ($null -ne $tmp) { $tmp.Dispose() }
            if ($null -ne $ms) { $ms.Dispose() }
        }
    }
    $Script:EmojiCache[$nom] = $img
    return $img
}

# Dibuixa l'emoji QUADRAT i centrat dins de $rect. Sense imatge, el text amb
# $lletra i $color (el comportament d'abans).
function _DibuixaEmoji($g, [string]$emoji, [System.Drawing.Rectangle]$rect, $lletra, [System.Drawing.Color]$color) {
    $img = _EmojiImatge $emoji
    if ($null -eq $img) {
        $fl = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::NoPadding
        [System.Windows.Forms.TextRenderer]::DrawText($g, $emoji, $lletra, $rect, $color, $fl)
        return
    }
    $m = [Math]::Min($rect.Width, $rect.Height)
    $dest = New-Object System.Drawing.Rectangle(($rect.X + [int](($rect.Width - $m) / 2)), ($rect.Y + [int](($rect.Height - $m) / 2)), $m, $m)
    $abans = $g.InterpolationMode
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    try { $g.DrawImage($img, $dest) } finally { $g.InterpolationMode = $abans }
}

function _BitmapIcona([string]$simbol, [single]$midaPt, [System.Drawing.Color]$color) {
    $f = New-Object System.Drawing.Font($Script:IconaLletra, $midaPt, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
    $fmt = [System.Drawing.StringFormat]::GenericTypographic
    try {
        # Si es un emoji amb imatge, la imatge EN COLOR, de l'alcada de la lletra.
        $emo = _EmojiImatge $simbol
        if ($null -ne $emo) {
            $h = [Math]::Max(8, [int]$f.Height)
            $bmp = New-Object System.Drawing.Bitmap($h, $h)
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            try {
                $g.Clear([System.Drawing.Color]::Transparent)
                _DibuixaEmoji $g $simbol (New-Object System.Drawing.Rectangle(1, 1, ($h - 2), ($h - 2))) $f $color
            } finally { $g.Dispose() }
            $bmp.Tag = 'icona'
            return $bmp
        }
        # Mida amb la MATEIXA API que dibuixa (GDI+): la de TextRenderer (GDI)
        # dona una amplada diferent i el simbol quedava retallat.
        $tmp = New-Object System.Drawing.Bitmap(1, 1)
        $gm = [System.Drawing.Graphics]::FromImage($tmp)
        try { $mida = $gm.MeasureString($simbol, $f, 400, $fmt) } finally { $gm.Dispose(); $tmp.Dispose() }
        $w = [int][Math]::Ceiling($mida.Width) + 2
        $h = [int][Math]::Ceiling($mida.Height) + 2
        $bmp = New-Object System.Drawing.Bitmap([Math]::Max(1, $w), [Math]::Max(1, $h))
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            # Fons TRANSPARENT i suavitzat en grisos (no ClearType): el boto
            # canvia de color en passar-hi el ratoli i la imatge l'ha de seguir.
            $g.Clear([System.Drawing.Color]::Transparent)
            $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
            $br = New-Object System.Drawing.SolidBrush($color)
            try { $g.DrawString($simbol, $f, $br, 1, 1, $fmt) } finally { $br.Dispose() }
        } finally { $g.Dispose() }
        $bmp.Tag = 'icona'
        return $bmp
    } finally { $f.Dispose() }
}

# Treu la icona que hi hagi posat _PosaIcona (un estat que ja no en porta).
function _TreuIcona($ctl) {
    $vella = $ctl.Image
    $ctl.Image = $null
    if ($null -ne $vella -and [string]$vella.Tag -eq 'icona') { $vella.Dispose() }
}

# Posa $simbol com a imatge al costat del text d'un boto (o casella) o d'una
# etiqueta. $text, si hi es, substitueix el text del control. -Despres, el
# simbol va al darrere (el triangle d'un desplegable). Crida-la DESPRES de
# posar el color del text: la icona el copia.
function _PosaIcona($ctl, [string]$simbol, $text = $null, [switch]$Despres) {
    _TreuIcona $ctl
    if ($null -ne $text) { $ctl.Text = [string]$text }
    $color = $ctl.ForeColor
    if ($color.IsEmpty) { $color = [System.Drawing.SystemColors]::ControlText }
    $img = _BitmapIcona $simbol ([single]$ctl.Font.SizeInPoints) $color
    if ($ctl -is [System.Windows.Forms.ButtonBase]) {
        $ctl.Image = $img
        $ctl.ImageAlign = $ctl.TextAlign
        $ctl.TextImageRelation = if ($Despres) { 'TextBeforeImage' } else { 'ImageBeforeText' }
        return
    }
    # Una ETIQUETA no sap posar la imatge al costat del text: la dibuixa a sobre.
    # Se li deixa lloc amb espais, que es l'unica cosa que no depen de la mida
    # de l'etiqueta (n'hi ha d'AutoSize i de mida fixa).
    $ctl.Image = $img
    $ctl.ImageAlign = if ($Despres) { 'TopRight' } else { 'TopLeft' }
    $espai = [Math]::Max(1, [System.Windows.Forms.TextRenderer]::MeasureText('a a', $ctl.Font).Width - [System.Windows.Forms.TextRenderer]::MeasureText('aa', $ctl.Font).Width)
    $n = [int][Math]::Ceiling(($img.Width + 3) / $espai)
    $buit = ([string][char]0x00A0) * $n
    $ctl.Text = if ($Despres) { [string]$ctl.Text + $buit } else { $buit + [string]$ctl.Text }
}

function _TxtEnrere  { return ([string][char]0x2190 + ' Enrere') }
function _TxtSeguent { return ('Seg' + [char]0x00FC + 'ent ' + [char]0x2192) }

# On va cada boto. PURA (es prova a Linux): rep l'amplada del contenidor i les
# de cada grup, i torna les X. El grup de la dreta es llegeix d'esquerra a
# dreta, com es veura: el darrer es el que queda enganxat al marge.
function _PeuPosicions([int]$ample, [int[]]$esquerra, [int[]]$dreta, [int]$marge = 15, [int]$sep = 10) {
    $xe = New-Object System.Collections.ArrayList
    $x = $marge
    foreach ($w in @($esquerra)) { [void]$xe.Add($x); $x += $w + $sep }
    $xd = New-Object System.Collections.ArrayList
    $x = $ample - $marge
    $amples = @($dreta)
    for ($i = $amples.Count - 1; $i -ge 0; $i--) { $x -= $amples[$i]; [void]$xd.Insert(0, $x); $x -= $sep }
    return @{ Esquerra = $xe.ToArray(); Dreta = $xd.ToArray() }
}

# L'amplada d'un boto amb aquest text: la que demana el text i, com a minim,
# 100 (que "OK" i "Tancar" no quedin esquifits). PURA si se li dona la mida
# del text.
function _PeuAmple([int]$ampleText, [int]$fix = 0) {
    if ($fix -gt 0) { return $fix }
    return [Math]::Max(100, $ampleText + 32)
}

function _AddPeuBotons($form, $esquerra, $dreta, [int]$y, $pare = $null, [switch]$Ancorat) {
    if ($null -eq $pare) { $pare = $form }
    $out = @{}
    $fets = @{ Left = (New-Object System.Collections.ArrayList); Right = (New-Object System.Collections.ArrayList) }
    foreach ($g in @(@{ Specs = @($esquerra); Costat = 'Left' }, @{ Specs = @($dreta); Costat = 'Right' })) {
        foreach ($s in @($g.Specs)) {
            if ($null -eq $s) { continue }
            $b = New-Object System.Windows.Forms.Button
            $b.Text = [string]$s.Text
            $estil = if ($s.Estil) { [string]$s.Estil } else { 'secundari' }
            if ($null -ne $Script:BrandMaroon) {
                switch ($estil) {
                    'primari' { _StylePrimaryButton $b }
                    'accent'  { _StyleAccentButton $b $s.Fons $s.FonsHover }
                    default   { _StyleSecondaryButton $b }
                }
            }
            $ampleText = [System.Windows.Forms.TextRenderer]::MeasureText($b.Text, $b.Font).Width
            if ($s.Icona -and $null -ne $Script:BrandMaroon) {
                _PosaIcona $b ([string]$s.Icona)
                $ampleText += $b.Image.Width + 4
            }
            $b.Size = New-Object System.Drawing.Size((_PeuAmple $ampleText ([int]$s.Ample)), 32)
            if ($s.Resultat) { $b.DialogResult = [string]$s.Resultat }
            if ($s.Clic) { $b.add_Click($s.Clic) }
            if ($s.Intro) { $form.AcceptButton = $b }
            if ($s.Esc) { $form.CancelButton = $b }
            $out[[string]$s.Nom] = $b
            [void]$fets[$g.Costat].Add($b)
        }
    }
    $amplesE = [int[]]@($fets.Left | ForEach-Object { [int]$_.Width })
    $amplesD = [int[]]@($fets.Right | ForEach-Object { [int]$_.Width })
    # ELS DE LA DRETA ES RECOLOQUEN A CADA CANVI DE MIDA, no amb l'Anchor.
    # L'ancoratge a la dreta es calcula contra l'amplada que el contenidor te
    # EN AQUELL MOMENT, i un panell de peu amb Dock (o una pestanya) encara no
    # te la bona quan s'hi posen els botons: fa 200 d'ample, i un boto posat a
    # x=900 s'hi quedaria a -700 de la vora per sempre. Recalcular-ho al Resize
    # no depen de quan s'ha fet el layout.
    $vert = if ($Ancorat) { 'Bottom' } else { 'Top' }
    $col = @{ Pare = $pare; Esq = @($fets.Left); Dre = @($fets.Right); AE = $amplesE; AD = $amplesD }
    $col.Posa = {
        $pos = _PeuPosicions ([int]$col.Pare.ClientSize.Width) $col.AE $col.AD
        for ($i = 0; $i -lt $col.Dre.Count; $i++) { $col.Dre[$i].Left = [int]$pos.Dreta[$i] }
        return $pos
    }.GetNewClosure()
    $pos = & $col.Posa
    for ($i = 0; $i -lt $col.Esq.Count; $i++) {
        $b = $col.Esq[$i]
        $b.Location = New-Object System.Drawing.Point([int]$pos.Esquerra[$i], $y)
        $b.Anchor = ($vert + ', Left')
        [void]$pare.Controls.Add($b)
    }
    for ($i = 0; $i -lt $col.Dre.Count; $i++) {
        $b = $col.Dre[$i]
        $b.Location = New-Object System.Drawing.Point([int]$pos.Dreta[$i], $y)
        $b.Anchor = ($vert + ', Left')
        [void]$pare.Controls.Add($b)
    }
    if ($col.Dre.Count -gt 0) { $pare.add_Resize({ [void](& $col.Posa) }.GetNewClosure()) }
    return $out
}
