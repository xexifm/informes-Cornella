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
#
# $ambHora = $false dona NOMES LA DATA. La fa servir la rajola "Copiar informes",
# que a l'espai de l'hora hi te l'interruptor A/M (l'hora, alli, no interessa).
function _FormatRunStamp([string]$iso, [bool]$ambHora = $true) {
    if ([string]::IsNullOrWhiteSpace($iso)) { return '(mai)' }
    $fmt = if ($ambHora) { 'dd/MM/yy HH:mm' } else { 'dd/MM/yy' }
    try { return ([datetime]::Parse($iso)).ToString($fmt) } catch { return '(mai)' }
}

# Llegeix una marca de temps d'"ultima execucio" d'un JSON d'estat. Torna el
# text ISO tal qual (o '' si no hi es): QUI la formata decideix si vol l'hora,
# i aixi el segell normal i el de l'interruptor llegeixen pel mateix cami.
function _LastRunIso($jsonPath, $prop) {
    if ([string]::IsNullOrWhiteSpace($jsonPath) -or -not (Test-Path -LiteralPath $jsonPath -ErrorAction SilentlyContinue)) { return '' }
    try {
        $o = Read-JsonFile $jsonPath
        if ($null -ne $o -and $o.PSObject.Properties[$prop]) {
            # Read-JsonIso (Json.ps1): el 5.1 deixa la marca com a cadena i el
            # pwsh 7 la converteix a [datetime], i un [datetime] a [string] surt
            # en el format de la CULTURA de la maquina -que _FormatRunStamp ja no
            # sap tornar a parsejar-. Normalitzar-ho es de tothom qui llegeix una
            # data d'un estat, per aixo es fa en UN sol lloc.
            return [string](Read-JsonIso $o.$prop)
        }
    } catch { }
    return ''
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
$Script:AccionsSenseSegell = @('nou', 'seguiment', 'actextr', 'llicencia', 'mnstrans', 'llicdb', 'config', 'editcataleg')

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

# La marca (ISO) del segell d'una eina, i el text que se'n pinta. Van partides
# perque l'interruptor de "Copiar informes" vol LA MATEIXA marca amb un altre
# format (data sola): duplicar la cerca hauria estat la manera que un dia els
# dos segells diguessin coses diferents.
function _LastRunIsoEina([string]$accio) {
    if ([string]::IsNullOrWhiteSpace($accio)) { return '' }
    if ($Script:SegellPropi.Contains($accio) -and -not [string]::IsNullOrWhiteSpace($LocalActivitatsDir)) {
        $d = $Script:SegellPropi[$accio]
        $t = _LastRunIso (Join-Path $LocalActivitatsDir $d.Fitxer) $d.Prop
        if (-not [string]::IsNullOrWhiteSpace($t)) { return $t }
    }
    return (_LastRunIso (_EinesStatePath) $accio)
}

function _LastRunEina([string]$accio) {
    return (_FormatRunStamp (_LastRunIsoEina $accio))
}

# ----------------------------------------------------------------------------
# L'AJUDA DE CADA EINA (el "?" petit de la cantonada de cada rajola)
# ----------------------------------------------------------------------------
# Una o dues frases que diuen QUE FA l'eina, indexades per ACCIO, com el segell.
# Clicar el "?" mostra el text i NO obre l'eina. Hi ha prova que CADA rajola del
# menu en te: una eina nova sense text fa petar les proves, no surt sense ajuda.
# Cometes dobles i apostrof recte: el tipografic tanca un literal amb '...'.
$Script:AjudaEines = @{
    ruta              = "Escrius els ID GIA de les activitats que vols visitar i et calcula la ruta més curta des de la base, amb un mapa numerat que pots imprimir."
    planol            = "Pinta les parcel·les de Cornellà on hi ha activitats, del color del seu estat (precintada, requeriment, sense informes, sense res pendent) i amb els ID GIA. Filtres per estat i locals buits."
    coordenades       = "Repassa per zones, sobre un mapa, on hi ha cada activitat i et deixa corregir-ne la posició. Et baixes un Excel amb les coordenades noves: no toca l'Excel d'activitats."
    precintades       = "Obre al navegador el mapa i el llistat públic de les activitats precintades."
    controlsperiodics = "Llista les activitats de l'annex II, III o de l'apartat 561 amb les dates dels controls periòdics (primer les que toquen abans). En pots generar els informes i els correus."
    recordatoris      = "Envia recordatoris periòdics per correu als titulars que tenen un requeriment o un precinte pendent, segons la base d'informes."
    informesdb        = "Recorre la carpeta dels informes fets i n'actualitza la base: la data, l'ID GIA i la conclusió de cada un. L'interruptor A/M de sota ho fa sol cada dia a les [HORA_AUTO]."
    informesdbedit    = "Mostra la base d'informes amb l'estat de cada activitat, amb filtres i exportació a CSV. Hi pots corregir la conclusió breu d'un informe o fer que s'ignori."
    copiarinformes    = "Copia els informes nous a la carpeta de còpia, tots junts, sense esborrar mai res. L'interruptor A/M de sota ho fa sol cada dia a les [HORA_AUTO]."
    convertirpdf      = "Converteix un informe de Word (o una carpeta sencera) a PDF i, si ho marques, el signa amb AutoFirma."
    comprovarexcel    = "Comprova que les activitats que la base d'informes té en Precinte / Cessament també ho tinguin marcat a l'Excel d'activitats, i et llista les que no."
    seguimentgia      = "Fa els llistats de seguiment de la base d'activitats (precintes, denúncies, requerits per decret, sonometria i annex II), en Excel o en PDF."
    emailtextos       = "Edita l'assumpte i el text del correu de requeriments que s'envia al titular, des del mòbil i des d'Enviar correu."
    enviarcorreu      = "Envia al titular, des de l'ordinador, el correu amb els requeriments d'un informe ja fet, amb el mateix format que el del mòbil."
    normativa         = "Baixa a la carpeta local\normativa el text vigent de totes les normes (les de REQ1 i les dels marcadors), cada una amb un nom que diu de quin tema és, i en fa un índex en Excel."
    revisio           = "Mira si la normativa dels requeriments encara és vigent, si els enllaços funcionen i si tots els punts de REQ1 tenen la fitxa d'informació. Pot baixar també la normativa nova. Ho deixa en un Excel."
    revisarmobil      = "Mira si han arribat informes preparats des del mòbil (per Google Drive) i en fa el Word."
}

function _AjudaEina([string]$accio) {
    if ([string]::IsNullOrWhiteSpace($accio) -or -not $Script:AjudaEines.Contains($accio)) { return '' }
    # [HORA_AUTO]: l'hora dels modes automatics (ModeAutomatic.ps1), que es
    # carrega DESPRES d'aquest fitxer; es resol en ensenyar el text.
    return ([string]$Script:AjudaEines[$accio]).Replace('[HORA_AUTO]', (Get-AutoHoraText))
}

# ON VA CADA GRUP D'EINES (pura: es prova sense WinForms).
#
# Abans les eines anaven SOTA els tipus d'informe, un grup per fila: la
# finestra feia ~980 px d'alt i en un portatil (o amb el Windows al 125%)
# calia fer scroll. Ara van en una COLUMNA a la dreta, en una GRAELLA de
# quatre rajoles d'ample (octubre 2026: l'usuari va demanar que tot quedes
# alineat, i amb files de 5, 4, 2+2 i 3 rajoles no hi havia graella). Els
# grups son de quatre (CARRER, TITULARS, BASE D'INFORMES) i els de dues
# comparteixen fila (GIA | NORMATIVA): com que entre grups hi ha el MATEIX
# espai que entre rajoles, NORMATIVA cau just a la tercera columna.
#
# Rep el nombre de rajoles de cada grup i les mides; torna, per a cada grup i
# en el mateix ordre, @{ X; Fila }: la X relativa a la columna i la fila (0, 1,
# ...). L'alcada de cada fila la decideix el menu, que les estira fins a quadrar
# amb el darrer boto d'informe. Un grup mai es parteix entre dues files, i un
# grup mes ample que la columna surt sol a la seva fila (no es perd).
function _MenuDisposaGrups([int[]]$rajoles, [int]$ampleMax, [int]$tileW, [int]$tileGap) {
    $out = New-Object System.Collections.ArrayList
    $x = 0; $fila = 0
    foreach ($n in @($rajoles)) {
        $w = ([Math]::Max(1, $n) * $tileW) + (([Math]::Max(1, $n) - 1) * $tileGap)
        if ($x -gt 0 -and ($x + $w) -gt $ampleMax) { $x = 0; $fila++ }
        [void]$out.Add(@{ X = $x; Fila = $fila })
        $x += $w + $tileGap
    }
    return ,($out.ToArray())
}

# On comenca cada fila de rajoles (pura). Les files s'ESTIREN perque la de baix
# acabi a la mateixa alcada que el darrer boto d'informe ($fi); si no hi caben
# amb el pas minim, el pas minim mana (i la columna d'eines sera la mes alta).
function _MenuFilesY([int]$files, [int]$y0, [int]$fi, [int]$altRajola, [int]$pasMinim) {
    if ($files -le 0) { return ,@() }
    $out = New-Object System.Collections.ArrayList
    if ($files -eq 1) { [void]$out.Add($y0); return ,($out.ToArray()) }
    $pas = [Math]::Max([double]$pasMinim, ([double]($fi - $y0 - $altRajola) / ($files - 1)))
    for ($i = 0; $i -lt $files; $i++) { [void]$out.Add([int]($y0 + [Math]::Round($i * $pas))) }
    return ,($out.ToArray())
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
    $icoMns  = [System.Char]::ConvertFromUtf32(0x1F504)  # fletxes: modificacio / transmissio
    $mnsNom  = ('Modificaci' + $oG + ' No Substancial / Transmissi' + $oG)

    # Menu ORDENAT. Cada entrada: Action, Label (nom amic), Sub (descripcio
    # curta en gris), Icon (emoji del xip), Doc (xip del document a la dreta) i,
    # per a 'nou', el Cataleg (FileInfo). Els 'nou' nomes surten si el .docx hi es.
    $menu = New-Object System.Collections.ArrayList
    # L'ORDRE DEL MENU EL DECIDEIX L'USUARI i es aquest (agost 2026):
    #   1 Requeriment - Nou   2 Requeriment - Seguiment   3 Llicencia
    #   4 Activitats extraordinaries   5 MNS / Transmissio   6 Ampliacio termini
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
    # MNS / TRANSMISSIO: ENTRADA PROPIA. Abans s'hi arribava des de dins de
    # Llicencia (el pas 1 oferia les cinc fases juntes) i no es veia des del
    # menu. Comparteixen capcalera, tramit i base de dades amb Llicencia -per
    # aixo passen pel mateix assistent-, pero son informes a part.
    [void]$menu.Add(@{ Action='mnstrans'; Label=$mnsNom; Sub='Informes curts'; Icon=$icoMns; Doc='MNSTRANS'; Cataleg=$null })
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
    # LA LINIA DELS TITOLS: el de l'esquerra (INFORMES) i els de la dreta
    # (CARRER...) son a la mateixa alcada i amb el mateix estil, i les rajoles
    # i els botons comencen tots a $yContingut. Abans "Que vols fer?" anava amb
    # una altra lletra i les rajoles comencaven 5 px mes amunt que els botons.
    $yTitols = $headerHeight + 23
    $yContingut = $yTitols + 22
    $fTitolGrup = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $colTitolGrup = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'INFORMES'
    $lbl.Font = $fTitolGrup
    $lbl.ForeColor = $colTitolGrup
    $lbl.Location = New-Object System.Drawing.Point(20, $yTitols)
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $fMain = New-Object System.Drawing.Font('Segoe UI', 12.5, [System.Drawing.FontStyle]::Bold)
    $fDet  = New-Object System.Drawing.Font('Segoe UI', 9.5,  [System.Drawing.FontStyle]::Regular)
    $fIcon = New-Object System.Drawing.Font('Segoe UI Emoji', 15, [System.Drawing.FontStyle]::Regular)
    $fEmoS = New-Object System.Drawing.Font('Segoe UI Emoji', 9, [System.Drawing.FontStyle]::Regular)
    $pencil = [string][char]0x270F + [char]0xFE0F   # emoji d'editar
    $emoXip = 16    # els emojis dels xips (llapis, Dades), en px
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
            # En COLOR (imatge): vegeu _DibuixaEmoji (UiFinestra.ps1).
            _DibuixaEmoji $g $ico (New-Object System.Drawing.Rectangle(($cx + 8), ($cy + 8), ($chip - 16), ($chip - 16))) $fIcon $colGranat
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
            $szP = New-Object System.Drawing.Size($emoXip, $emoXip)
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
            _DibuixaEmoji $g $pencil (New-Object System.Drawing.Rectangle(($dx + $pad), ($dy + [int](($chH - $emoXip) / 2)), $emoXip, $emoXip)) $fEmoS $colGranat
            [System.Windows.Forms.TextRenderer]::DrawText($g, $doc, $fDet, (New-Object System.Drawing.Point(($dx + $pad + $szP.Width + $gap), ($dy + 4))), $colGranat, $flags)
            $entry.DocChipRect = New-Object System.Drawing.Rectangle($dx, $dy, $cw, $chH)

            # Xip EXTRA (opcional), just a l'esquerra del del document. Mateix
            # aspecte i mateix hit-test; el seu rectangle va a $entry.ExtraChipRect.
            if ($null -ne $entry.Extra) {
                $et = [string]$entry.Extra.Text
                $ei = [string]$entry.Extra.Icon
                $szEI = New-Object System.Drawing.Size($emoXip, $emoXip)
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
                _DibuixaEmoji $g $ei (New-Object System.Drawing.Rectangle(($ex + $pad), ($dy + [int](($chH - $emoXip) / 2)), $emoXip, $emoXip)) $fEmoS $colGranat
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
    # ALINEATS A LA DRETA de la columna d'informes (el marge dret dels botons),
    # a la linia dels titols: es posen despres, quan ja se'n sap l'amplada.
    $llsComuns = New-Object System.Collections.ArrayList
    foreach ($d in @(
        @{ Doc = '0 CAPCALERA';   Text = ('Cap' + [char]0x00E7 + 'alera') },
        @{ Doc = '0 CONCLUSIONS'; Text = 'Conclusions' })) {
        $ll = New-Object System.Windows.Forms.LinkLabel
        $ll.Text = [string]$d.Text
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
        [void]$llsComuns.Add($ll)
    }
    $xComuns = 20 + 560
    for ($il = $llsComuns.Count - 1; $il -ge 0; $il--) {
        $xComuns -= $llsComuns[$il].PreferredWidth
        $llsComuns[$il].Location = New-Object System.Drawing.Point($xComuns, ($yTitols - 1))
        $xComuns -= 14
    }
    $y = $yContingut
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

    # ---- Eines: en una COLUMNA a la dreta dels tipus d'informe ---------------
    # (vegeu _MenuDisposaGrups: a sota no hi cabien sense fer scroll)
    $fiInformes = $y - 8     # on acaba el darrer boto (pas de 70, boto de 62)

    # LES EINES, PER MOMENT DE LA FEINA (octubre 2026, amb l'usuari): abans
    # s'agrupaven per d'on treuen les dades i sortien files de 5, 4, 2+2 i 3;
    # ara son quatre grups de quatre i fan una graella. Comportament per
    # rajola: 'action' tanca el menu amb l'accio; 'url' obre l'enllac public
    # SENSE tancar el menu (precintades).
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
    $tiSobre = [System.Char]::ConvertFromUtf32(0x2709)    # ✉
    $tiList  = [System.Char]::ConvertFromUtf32(0x1F4CA)   # 📊
    $tiMap   = [System.Char]::ConvertFromUtf32(0x1F5FA)   # 🗺
    $tiBell  = [System.Char]::ConvertFromUtf32(0x1F514)   # 🔔
    $tiLlibres = [System.Char]::ConvertFromUtf32(0x1F4DA) # 📚
    $tiLupa    = [System.Char]::ConvertFromUtf32(0x1F50D) # 🔍
    $tiCases   = [System.Char]::ConvertFromUtf32(0x1F3D8) # 🏘
    # CARRER: preparar la inspeccio i el que es porta del carrer amb el mobil.
    $carrer = @(
        @{ Emoji = $tiCases; Label = ('Pl' + [char]0x00E0 + 'nol activitats'); Kind = 'action'; Action = 'planol' }
        @{ Emoji = $tiPin;   Label = 'Generar ruta';           Kind = 'action'; Action = 'ruta' }
        @{ Emoji = $tiMap;   Label = 'Coordenades';            Kind = 'action'; Action = 'coordenades' }
        # 'Action' tambe a la rajola d'enllac: no despatxa res, pero es la clau
        # del seu segell d'ultima execucio.
        @{ Emoji = $tiLock;  Label = 'Activitats precintades'; Kind = 'url';    Action = 'precintades'; Url = $urlPrec }
        @{ Emoji = $tiInbox; Label = ('Revisar m' + [char]0x00F2 + 'bil'); Kind = 'action'; Action = 'revisarmobil' }
    )
    # TITULARS: el que s'envia o es reclama als titulars.
    $titulars = @(
        @{ Emoji = $tiMail;  Label = 'Enviar correu';     Kind = 'action'; Action = 'enviarcorreu' }
        @{ Emoji = $tiSobre; Label = 'Textos del correu'; Kind = 'action'; Action = 'emailtextos' }
        @{ Emoji = $tiBell;  Label = 'Recordatoris';      Kind = 'action'; Action = 'recordatoris' }
        @{ Emoji = $tiCal;   Label = ('Controls peri' + [char]0x00F2 + 'dics'); Kind = 'action'; Action = 'controlsperiodics' }
    )
    # BASE D'INFORMES: eines de la base d'informes + conversio a PDF.
    $reports = @(
        # Les rajoles amb INTERRUPTOR: sota seu, on les altres tenen l'hora, hi
        # va el commutador A/M del mode automatic (vegeu mes avall i
        # $Script:ModesAuto a ModeAutomatic.ps1).
        @{ Emoji = $tiBox;   Label = 'Actualitzar base'; Kind = 'action'; Action = 'informesdb'; Interruptor = $true }
        @{ Emoji = $tiClip;  Label = 'Editar base';      Kind = 'action'; Action = 'informesdbedit' }
        @{ Emoji = $tiCopy;  Label = 'Copiar informes';  Kind = 'action'; Action = 'copiarinformes'; Interruptor = $true }
        @{ Emoji = $tiPdf;   Label = 'Word a PDF';       Kind = 'action'; Action = 'convertirpdf' }
    )
    # GIA: eines que parlen de la base de dades d'ACTIVITATS (el GIA), no dels
    # informes. Comparteix fila amb NORMATIVA.
    $gia = @(
        @{ Emoji = $tiList;  Label = 'Seguiment';       Kind = 'action'; Action = 'seguimentgia' }
        @{ Emoji = $tiCheck; Label = 'Comprovar Excel'; Kind = 'action'; Action = 'comprovarexcel' }
    )
    # NORMATIVA: la normativa dels requeriments i mantenir-los al dia.
    $normativaRow = @(
        @{ Emoji = $tiLlibres; Label = 'Normativa'; Kind = 'action'; Action = 'normativa' }
        @{ Emoji = $tiLupa;    Label = 'Revisar requeriments'; Kind = 'action'; Action = 'revisio' }
    )
    $fTileIco   = New-Object System.Drawing.Font('Segoe UI Emoji', 14, [System.Drawing.FontStyle]::Regular)
    $fTileTxt   = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Regular)
    $tileBorder = [System.Drawing.Color]::FromArgb(214, 219, 225)
    $tileTxtCol = [System.Drawing.Color]::FromArgb(63, 73, 85)
    $fAjuda        = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
    $colAjuda      = [System.Drawing.Color]::FromArgb(247, 231, 234)
    $colAjudaHover = [System.Drawing.Color]::FromArgb(166, 26, 47)
    $tilePaint = {
        param($s, $e)
        $t = $s.Tag
        $g = $e.Graphics
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $rc = $s.ClientRectangle
        $flW = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::WordBreak -bor [System.Windows.Forms.TextFormatFlags]::NoPadding
        # L'emoji EN COLOR (imatge, vegeu _DibuixaEmoji) i l'etiqueta a sota.
        _DibuixaEmoji $g ([string]$t.Emoji) (New-Object System.Drawing.Rectangle(0, 6, $rc.Width, 22)) $fTileIco ([System.Drawing.Color]::Black)
        $lbRect = New-Object System.Drawing.Rectangle(4, 31, ($rc.Width - 8), ($rc.Height - 33))
        [System.Windows.Forms.TextRenderer]::DrawText($g, $t.Label, $fTileTxt, $lbRect, $tileTxtCol, $flW)
        # El "?" de l'ajuda, a la cantonada de dalt a la dreta: rodona granat
        # suau, i granat plena quan hi passa el ratoli (com els xips de dalt).
        # Un "?" normal, no un emoji: amb la Segoe UI surt sempre.
        $t.AjudaRect = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$t.Ajuda)) {
            $hr = New-Object System.Drawing.Rectangle(($rc.Width - 17), 3, 14, 14)
            $bH = New-Object System.Drawing.SolidBrush($(if ($t.AjudaHover) { $colAjudaHover } else { $colAjuda }))
            $g.FillEllipse($bH, $hr)
            $bH.Dispose()
            $flH = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::NoPadding
            $colQ = if ($t.AjudaHover) { [System.Drawing.Color]::White } else { $colAjudaHover }
            [System.Windows.Forms.TextRenderer]::DrawText($g, '?', $fAjuda, $hr, $colQ, $flH)
            # Una mica mes gran que el dibuix: 14 pixels son poc per encertar-hi.
            $t.AjudaRect = New-Object System.Drawing.Rectangle(($hr.X - 3), 0, ($hr.Width + 6), ($hr.Height + 6))
        }
    }.GetNewClosure()
    $tileClick = {
        param($s, $e)
        $t = $s.Tag
        # Clic al "?": nomes l'explicacio, l'eina NO s'obre. El Click no porta
        # coordenades; les donen la posicio del ratoli i el rectangle que ha
        # guardat el Paint.
        $pos = $s.PointToClient([System.Windows.Forms.Control]::MousePosition)
        if ($null -ne $t.AjudaRect -and $t.AjudaRect.Contains($pos)) {
            [System.Windows.Forms.MessageBox]::Show([string]$t.Ajuda, [string]$t.Label, 'OK', 'Information') | Out-Null
            return
        }
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
    # El "?" es ressalta i el cursor passa a "ma" quan el ratoli hi es a sobre
    # (nomes es repinta quan canvia, com el xip de l'editor de catalegs).
    $tileMove = {
        param($s, $e)
        $t = $s.Tag
        $sobre = ($null -ne $t.AjudaRect -and $t.AjudaRect.Contains($e.Location))
        if ($sobre -ne [bool]$t.AjudaHover) {
            $t.AjudaHover = $sobre
            $s.Cursor = if ($sobre) { [System.Windows.Forms.Cursors]::Hand } else { [System.Windows.Forms.Cursors]::Default }
            $s.Invalidate()
        }
    }.GetNewClosure()
    $tileLeave = {
        param($s, $e)
        $t = $s.Tag
        if ([bool]$t.AjudaHover) { $t.AjudaHover = $false; $s.Cursor = [System.Windows.Forms.Cursors]::Default; $s.Invalidate() }
    }.GetNewClosure()
    # CINC columnes (octubre 2026: CARRER en te cinc, amb el Planol activitats).
    # 84 + 8: deixa 76 px de text, on encara hi caben 'Activitats precintades',
    # 'Revisar requeriments' i 'Controls periodics' en dues linies sense trencar
    # cap paraula ('requeriments', la mes llarga, fa ~64 px a Segoe UI 8).
    # Totes les files igual d'estretes, perque la graella segueixi alineada.
    $tileW = 84; $tileH = 62; $tileGap = 8
    $Script:MenuColumnes = 5
    # Sota CADA rajola, en petit, l'ultima vegada que s'ha fet servir l'eina
    # ('(mai)' si encara no). El segell es llegeix per l'ACCIO de la rajola (la
    # clau del registre), no per la posicio dins de la fila: abans els indexs
    # anaven a pinyo fix contra una fila concreta i, en moure 'Comprovar Excel'
    # de fila, el segell hauria anat a la rajola equivocada.
    $fStamp = New-Object System.Drawing.Font('Segoe UI', 7)
    $colStamp = [System.Drawing.Color]::FromArgb(120, 128, 138)
    $ttEines = New-Object System.Windows.Forms.ToolTip
    # ------------------------------------------------------------------------
    # L'INTERRUPTOR A/M del mode automatic ("Copiar informes", "Actualitzar base")
    # ------------------------------------------------------------------------
    # Cada eina que en te s'apunta a $Script:ModesAuto (ModeAutomatic.ps1) i la
    # seva rajola porta 'Interruptor = $true'; el menu no sap res mes de cap.
    #
    # Va A L'ESPAI DEL SEGELL d'aquella rajola: la data on hi havia la data i el
    # commutador on hi havia l'hora, que es el que l'usuari va demanar (l'hora
    # de l'ultima copia no li interessa). No ocupa ni un pixel mes que les
    # altres rajoles.
    #
    #   A (verd)  mode automatic: cada dia a la seva hora amb el programa obert
    #             i, si aquell venciment no s'ha servit, en obrir el programa.
    #             Es fa en segon pla i no s'hi veu res (SiToca del registre).
    #   M (gris)  mode manual: nomes es fa quan cliques la rajola.
    #
    # I LA DATA DIU QUI VA FER L'ULTIMA PASSADA: verda si la va fer el mode
    # automatic, grisa si la vas fer tu. Aixi, d'un cop d'ull, se sap si
    # l'automatic esta treballant de debo o nomes esta ences.
    #
    # LA RAJOLA SEGUEIX FUNCIONANT SEMPRE, digui el que digui l'interruptor: el
    # commutador es un control a part (el clic es mira contra el seu rectangle),
    # o sigui que clicar la rajola mai el toca ni al reves.
    $colAuto     = [System.Drawing.Color]::FromArgb(46, 160, 67)    # verd de la pastilla
    $colAutoText = [System.Drawing.Color]::FromArgb(26, 122, 55)    # el mateix verd, mes fosc: a 7pt el clar no es llegeix
    $colManual   = [System.Drawing.Color]::FromArgb(150, 155, 163)  # gris de la pastilla
    $fSwitch = New-Object System.Drawing.Font('Segoe UI', 6, [System.Drawing.FontStyle]::Bold)

    # L'estat que pinta cada interruptor, en UN hashtable PER RAJOLA, al Tag del
    # seu Panel: aixi el rellotge i el clic el refresquen sense tornar a muntar
    # cap control, i els mateixos scriptblocks serveixen per a totes les
    # rajoles. $autos: tots, per al rellotge.
    #   @{ Accio; Mode (l'entrada del registre); On; Data; Verd; Ctl; Rect; Tip }
    # El registre es CAPTURA en una variable: dins d'una closure, $Script: no es
    # el de l'script (vegeu CLAUDE.md).
    $modesAuto = $Script:ModesAuto
    $autos = New-Object System.Collections.ArrayList
    $refrescaAuto = {
        param($auto)
        $m = $auto.Mode
        $auto.On   = [bool](& $m.Actiu)
        $auto.Verd = ([string](& $m.UltimMode) -eq 'auto')
        $auto.Data = [string](_FormatRunStamp (_LastRunIsoEina ([string]$auto.Accio)) $false)
        if ($null -ne $auto.Ctl) {
            $auto.Ctl.Invalidate()
            if ($null -ne $auto.Tip) {
                $q = if ($auto.On) { [string]$m.TipA } else { [string]$m.TipM }
                $auto.Tip.SetToolTip($auto.Ctl, $q)
            }
        }
    }.GetNewClosure()

    # La pastilla es dibuixa amb DOS SEMICERCLES I UN RECTANGLE: el GDI+ no te
    # rectangle arrodonit i muntar-ne un amb GraphicsPath serien vint linies mes
    # per a 24x12 pixels.
    $autoPaint = {
        param($s, $e)
        $auto = $s.Tag
        if ($null -eq $auto) { return }
        $g = $e.Graphics
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $rc = $s.ClientRectangle
        $flN = [System.Windows.Forms.TextFormatFlags]::NoPadding
        $flC = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor $flN
        $flV = [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor $flN
        $pw = 24; $ph = 12; $gap = 5
        $txt = [string]$auto.Data
        $szT = [System.Windows.Forms.TextRenderer]::MeasureText($g, $txt, $fStamp, [System.Drawing.Size]::Empty, $flN)
        $x0 = [int](($rc.Width - ($szT.Width + $gap + $pw)) / 2)
        if ($x0 -lt 0) { $x0 = 0 }

        # La data, verda si l'ultima passada la va fer el mode automatic.
        $colTxt = if ($auto.Verd) { $colAutoText } else { $colStamp }
        $rT = New-Object System.Drawing.Rectangle($x0, 0, $szT.Width, $rc.Height)
        [System.Windows.Forms.TextRenderer]::DrawText($g, $txt, $fStamp, $rT, $colTxt, $flV)

        $px = $x0 + $szT.Width + $gap
        $py = [int](($rc.Height - $ph) / 2)
        $col = if ($auto.On) { $colAuto } else { $colManual }
        $br = New-Object System.Drawing.SolidBrush($col)
        $g.FillEllipse($br, $px, $py, $ph, $ph)
        $g.FillEllipse($br, ($px + $pw - $ph), $py, $ph, $ph)
        $g.FillRectangle($br, ($px + [int]($ph / 2)), $py, ($pw - $ph), $ph)
        $br.Dispose()

        # El boto blanc al costat que toca i la lletra a l'altre: A a l'esquerra
        # amb el boto a la dreta (ences), M a la dreta amb el boto a l'esquerra.
        $kn = $ph - 4
        $kx = if ($auto.On) { $px + $pw - $ph + 2 } else { $px + 2 }
        $bw = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $g.FillEllipse($bw, $kx, ($py + 2), $kn, $kn)
        $bw.Dispose()
        $lletra = if ($auto.On) { 'A' } else { 'M' }
        $lx = if ($auto.On) { $px } else { $px + $pw - $ph }
        $rLl = New-Object System.Drawing.Rectangle($lx, $py, $ph, $ph)
        [System.Windows.Forms.TextRenderer]::DrawText($g, $lletra, $fSwitch, $rLl, [System.Drawing.Color]::White, $flC)

        # El rectangle del clic el guarda el PAINT (com els xips de les rajoles):
        # es l'unic lloc que sap on ha quedat la pastilla despres de centrar-ho
        # tot segons l'ample que ocupi la data.
        $auto.Rect = New-Object System.Drawing.Rectangle(($px - 3), 0, ($pw + 6), $rc.Height)
    }.GetNewClosure()

    $autoClick = {
        param($s, $e)
        $auto = $s.Tag
        if ($null -eq $auto -or $null -eq $auto.Rect -or -not $auto.Rect.Contains($e.Location)) { return }
        $nou = -not $auto.On
        $m = $auto.Mode
        # Si li falta la configuracio (la carpeta de copia, la d'informes),
        # deixar-ho ences seria un automatic que no fa res i no ho diu.
        if ($nou) {
            $req = [string](& $m.Requisit)
            if ($req -ne '') {
                [System.Windows.Forms.MessageBox]::Show($req, [string]$m.Titol, 'OK', 'Information') | Out-Null
                return
            }
        }
        [void](& $m.DesaActiu $nou)
        & $refrescaAuto $auto
        # En engegar-lo, si el venciment d'avui ja ha passat i ningu no l'ha
        # servit, la passada surt ARA (esperar a dema no seria "automatic").
        if ($nou) { [void](& $m.SiToca) }
    }.GetNewClosure()

    # Feedback de que es clicable, igual que el xip de l'editor de catalegs.
    $autoMove = {
        param($s, $e)
        $auto = $s.Tag
        $sobre = ($null -ne $auto -and $null -ne $auto.Rect -and $auto.Rect.Contains($e.Location))
        $c = if ($sobre) { [System.Windows.Forms.Cursors]::Hand } else { [System.Windows.Forms.Cursors]::Default }
        if ($s.Cursor -ne $c) { $s.Cursor = $c }
    }.GetNewClosure()

    # Dibuixa les rajoles d'un grup amb el seu segell a partir de ($xRow, $yRow)
    # i retorna la $y de sota (helper unic: el fan servir tots els grups).
    $addTileRow = {
        param($items, $xRow, $yRow)
        $tx = $xRow
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
            $tool.Ajuda = _AjudaEina ([string]$tool.Action)
            $tb.add_Paint($tilePaint)
            $tb.add_Click($tileClick)
            $tb.add_MouseMove($tileMove)
            $tb.add_MouseLeave($tileLeave)
            [void]$form.Controls.Add($tb)

            if ([bool]$tool.Interruptor -and $modesAuto.Contains([string]$tool.Action)) {
                # MATEIX ESPAI que el segell de les altres: data + interruptor
                # A/M alla on elles tenen data + hora. Es un Panel dibuixat a ma
                # (un Label no pot portar la pastilla) i el clic es mira contra
                # el rectangle del commutador, no contra tot el control.
                $pnS = New-Object System.Windows.Forms.Panel
                $pnS.Location = New-Object System.Drawing.Point($tx, ($yRow + $tileH + 1))
                $pnS.Size = New-Object System.Drawing.Size($tileW, 15)
                $pnS.BackColor = $form.BackColor
                $auto = @{ Accio = [string]$tool.Action; Mode = $modesAuto[[string]$tool.Action]
                           On = $false; Data = '(mai)'; Verd = $false; Ctl = $pnS; Rect = $null; Tip = $ttEines }
                $pnS.Tag = $auto
                [void]$autos.Add($auto)
                $pnS.add_Paint($autoPaint)
                $pnS.add_MouseClick($autoClick)
                $pnS.add_MouseMove($autoMove)
                [void]$form.Controls.Add($pnS)
                & $refrescaAuto $auto
            } else {
                $lblS = New-Object System.Windows.Forms.Label
                $lblS.Text = [string](_LastRunEina ([string]$tool.Action))
                $lblS.Font = $fStamp
                $lblS.ForeColor = $colStamp
                $lblS.TextAlign = 'MiddleCenter'
                $lblS.Location = New-Object System.Drawing.Point($tx, ($yRow + $tileH + 2))
                $lblS.Size = New-Object System.Drawing.Size($tileW, 14)
                [void]$form.Controls.Add($lblS)
                # El guardem a la propia rajola: la d'enllac (precintades) no tanca
                # el menu i s'ha de poder refrescar el seu segell alli mateix.
                $tool.StampLabel = $lblS
            }

            $tx += $tileW + $tileGap
        }
        return ($yRow + $tileH + 20)
    }.GetNewClosure()
    $grups = @(
        @{ Titol = 'CARRER';    Items = $carrer }
        @{ Titol = 'TITULARS';  Items = $titulars }
        @{ Titol = ('BASE D' + [char]39 + 'INFORMES'); Items = $reports }
        @{ Titol = 'GIA';       Items = $gia }
        @{ Titol = 'NORMATIVA'; Items = $normativaRow }
    )
    # LA GRAELLA: cinc columnes a la dreta dels botons, separades per una
    # ratlla. La primera fila de rajoles comenca a $yContingut, com el primer
    # boto, i les files s'estiren fins que la de baix -amb el seu segell- acaba
    # on acaba el darrer boto (_MenuFilesY). Els titols van 22 px per sobre de
    # cada fila, com el d'INFORMES.
    $xEines = 20 + 560 + 40
    $ampleEines = ($Script:MenuColumnes * $tileW) + (($Script:MenuColumnes - 1) * $tileGap)
    $altSegell = 16
    $posGrups = _MenuDisposaGrups ([int[]]@($grups | ForEach-Object { @($_.Items).Count })) $ampleEines $tileW $tileGap
    $nFiles = 1 + (($posGrups | ForEach-Object { [int]$_.Fila } | Measure-Object -Maximum).Maximum)
    $filesY = _MenuFilesY $nFiles $yContingut $fiInformes ($tileH + $altSegell) ($tileH + $altSegell + 30)
    $yFiEines = $yContingut
    for ($ig = 0; $ig -lt $grups.Count; $ig++) {
        $gx = $xEines + [int]$posGrups[$ig].X
        $gy = [int]$filesY[[int]$posGrups[$ig].Fila]
        $sep = New-Object System.Windows.Forms.Label
        $sep.Text = [string]$grups[$ig].Titol
        $sep.Font = $fTitolGrup
        $sep.ForeColor = $colTitolGrup
        $sep.Location = New-Object System.Drawing.Point($gx, ($gy - 22))
        $sep.AutoSize = $true
        [void]$form.Controls.Add($sep)
        [void](& $addTileRow $grups[$ig].Items $gx $gy)
        $yFiEines = [Math]::Max($yFiEines, ($gy + $tileH + $altSegell))
        # Un grup que comparteix fila (NORMATIVA) porta una ratlla fina al
        # davant, al mig de l'espai entre columnes.
        if ([int]$posGrups[$ig].X -gt 0) {
            $rg = New-Object System.Windows.Forms.Label
            $rg.BackColor = [System.Drawing.Color]::FromArgb(230, 233, 237)
            $rg.Location = New-Object System.Drawing.Point(($gx - [int]($tileGap / 2) - 1), ($gy - 20))
            $rg.Size = New-Object System.Drawing.Size(1, ($tileH + 18))
            [void]$form.Controls.Add($rg)
        }
    }
    $yFi = [Math]::Max($fiInformes, $yFiEines)

    # Una ratlla fina separa les dues columnes, de la linia dels titols a baix.
    $ratlla = New-Object System.Windows.Forms.Label
    $ratlla.BackColor = $tileBorder
    $ratlla.Location = New-Object System.Drawing.Point((20 + 560 + 20), $yTitols)
    $ratlla.Size = New-Object System.Drawing.Size(1, ($yFi - $yTitols))
    [void]$form.Controls.Add($ratlla)
    $y = $yFi + 8

    # (Configuracio i Ajuda ja no son botons grans: van DISCRETS a la cantonada
    #  de la banda granat, mes avall.)
    $urlAjuda = 'https://github.com/xexifm/informes-cornella/blob/main/LLEGEIX-ME.md'

    $form.ClientSize = New-Object System.Drawing.Size(($xEines + $ampleEines + 20), ($y + 12))

    # Banda de capcalera GRANAT amb escut blanc (helper comu del redisseny).
    # S'afegeix al final (Dock=Top) per no desplacar els controls ja posicionats.
    $subTitle = 'Ajuntament de Cornell' + [char]0x00E0 + ' de Llobregat'
    $band = _AddBrandHeader $form "Generador d'informes" $subTitle $headerHeight

    # Botons DISCRETS a la cantonada dreta de la banda: Ajuda (?) i Configuracio
    # (rosca). Fons granat una mica mes clar, text blanc, sense vora. Ancorats a
    # la dreta perque segueixin la cantonada si es maximitza.
    # Els botons de la banda acaben al MATEIX marge dret (20) que les rajoles.
    $wForm = $form.ClientSize.Width
    $fBandIco = New-Object System.Drawing.Font('Segoe UI Emoji', 11, [System.Drawing.FontStyle]::Regular)
    $btnAjuda = New-Object System.Windows.Forms.Button
    $btnAjuda.Text = [string][char]0x2753
    $btnAjuda.Font = $fBandIco
    $btnAjuda.Size = New-Object System.Drawing.Size(30, 30)
    $btnAjuda.Location = New-Object System.Drawing.Point(($wForm - 50), 13)
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
    $btnCarpeta.Location = New-Object System.Drawing.Point(($wForm - 126), 13)
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
    $btnConfig.Location = New-Object System.Drawing.Point(($wForm - 88), 13)
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

    # ACTUALITZAR, amb text i no nomes la icona: l'usuari el fa servir molt i
    # abans era a Configuracio (dos clics endins). Fa EXACTAMENT el mateix que
    # aquell boto (Invoke-ActualitzarPrograma, Configuracio.ps1): demana
    # confirmacio, llanca Actualitzar.bat i tanca el programa. La fletxa en
    # cercle no es a la Segoe UI: va amb _PosaIcona.
    $btnActualitzarM = New-Object System.Windows.Forms.Button
    $btnActualitzarM.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Regular)
    $btnActualitzarM.FlatStyle = 'Flat'
    $btnActualitzarM.ForeColor = [System.Drawing.Color]::White
    $btnActualitzarM.BackColor = [System.Drawing.Color]::FromArgb(150, 45, 60)
    $btnActualitzarM.FlatAppearance.BorderSize = 0
    $btnActualitzarM.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $btnActualitzarM.TextAlign = 'MiddleCenter'
    _PosaIcona $btnActualitzarM ([string][char]0x21BB) ' Actualitzar'
    $btnActualitzarM.Size = New-Object System.Drawing.Size(118, 30)
    $btnActualitzarM.Location = New-Object System.Drawing.Point(($btnCarpeta.Left - 8 - 118), 13)
    $btnActualitzarM.Anchor = 'Top,Right'
    $btnActualitzarM.add_Click({ Invoke-ActualitzarPrograma })
    [void]$band.Controls.Add($btnActualitzarM)
    $ttBand.SetToolTip($btnActualitzarM, 'Baixa la versio nova del programa (Actualitzar.bat) i el torna a obrir')

    # ------------------------------------------------------------------------
    # EL RELLOTGE dels modes automatics ("Copiar informes", "Actualitzar base")
    # ------------------------------------------------------------------------
    # Un Timer de WinForms i no un bucle: el menu ha de seguir responent. Cada
    # minut demana a cada eina del registre (SiToca) si toca la passada -ella ho
    # decideix tot: si l'interruptor esta ences i si el venciment de la seva
    # hora encara no s'ha servit- i despres refresca els segells, que es l'unica
    # cosa que es veu quan la passada ja s'ha fet.
    #
    # LA PRIMERA COMPROVACIO ES AL 'Shown', no aqui: es la de "en obrir el
    # programa". Com que el menu es torna a obrir a cada volta de Main, tambe
    # es mira en tornar de qualsevol eina; repetir-ho no costa res perque la
    # marca 'auto_el' ja diu que aquell venciment esta servit.
    $tmrAuto = New-Object System.Windows.Forms.Timer
    $tmrAuto.Interval = 60000
    $tmrAuto.add_Tick({
        foreach ($m in @($modesAuto.Values)) { [void](& $m.SiToca) }
        foreach ($a in @($autos)) { & $refrescaAuto $a }
    }.GetNewClosure())
    $form.add_Shown({
        foreach ($m in @($modesAuto.Values)) { [void](& $m.SiToca) }
        $tmrAuto.Start()
    }.GetNewClosure())
    # El rellotge MOR AMB LA FINESTRA: un Timer viu que dispari sobre controls
    # ja destruits es una excepcio dins del bucle de missatges.
    $form.add_FormClosed({ try { $tmrAuto.Stop(); $tmrAuto.Dispose() } catch { } }.GetNewClosure())

    $res = $form.ShowDialog()
    if ($res -ne 'OK' -or $null -eq $result.Choice) { exit 0 }
    $ch = $result.Choice
    return @{ Action = $ch.Action; Cataleg = $ch.Cataleg; Doc = $ch.Doc }
}
