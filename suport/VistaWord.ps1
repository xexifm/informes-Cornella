#requires -Version 5.1
<#
.SYNOPSIS
  Genera una VISTA en Word (.docx) de cada cataleg, a partir del seu JSON.

.DESCRIPTION
  La FONT DE VERITAT del programa son els JSON d'ESTRUCTURALS. Els .docx ja no
  serveixen per generar res (l'unica excepcio es '0 CAPCALERA.docx', que SI que
  es una plantilla de veritat: la generacio en copia el fitxer i hi substitueix
  els <<PLACEHOLDERS>>).

  Aquest modul escriu, per a cada cataleg, un .docx amb TOT el contingut possible
  (tots els requeriments, totes les conclusions...) perque es pugui consultar
  comodament sense obrir el programa. Es una VISTA de sortida (el text tal com
  sortiria a l'informe) pero amb els TITOLS DE WORD posats, de manera que el
  panell de navegacio de Word mostri l'estructura:

    Titol 1  -> seccio del cataleg (o grup de conclusions = tipus d'informe)
    Titol 2  -> subseccio, o be l'item quan no hi ha subseccio
    Titol 3  -> item dins d'una subseccio
    Normal   -> el cos, amb la negreta/cursiva i els enllacos

  Es regenera automaticament en desar des de l'editor de catalegs i des de
  Actualitzar.bat, sobreescrivint el .docx antic del mateix nom.

  Les funcions de text son PURES (testejables en headless); nomes l'escriptura
  necessita Word (COM), i per tant nomes va a Windows.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES (testejables)
# ----------------------------------------------------------------------------

# Ruta del .docx de la vista d'un cataleg: MATEIX NOM, pero a
# local\vistes-catalegs\ (i no al costat del .json). Les vistes son DERIVADES:
# es regeneren soles i no es pugen. Aixi ESTRUCTURALS es queda nomes amb les
# FONTS (els .json + '0 CAPCALERA.docx') i no es barregen font i derivat.
#
# $dir permet dir on han d'anar (les proves hi passen una carpeta temporal);
# buit -> local\vistes-catalegs\ del clone.
# El [string] del 'return' NO es decoratiu: Join-Path es un cmdlet i el que en
# surt ve embolcallat en un PSObject, i despres $doc.SaveAs([ref]$out) peta amb
# "no se puede convertir el valor ... de tipo psobject al tipo Object".
function _VistaWordPathFor([string]$jsonPath, [string]$dir = '') {
    $nom = [System.IO.Path]::GetFileNameWithoutExtension([string]$jsonPath) + '.docx'
    $d = if ([string]::IsNullOrWhiteSpace($dir)) { Get-LocalSubdir $RepoRoot 'Vistes' } else { $dir }
    return [string](Join-Path $d $nom)
}

# '0 CAPCALERA' es una plantilla de VERITAT: no se n'ha de generar mai cap vista
# (la sobreescriuriem i perdriem la carta amb l'escut i la taula).
#
# LLIC tampoc en te: no es un cataleg de deficiencies sino la capa propia de
# Llicencia sobre REQ1 (per cada requeriment, el "No es disposa", el "Es
# disposa" i el "Quan:"). Els seus items no porten text -el treuen de REQ1 en
# viu-, o sigui que la vista sortiria plena de punts buits.
function _VistaEsProtegit([string]$jsonPath) {
    # NOMES la plantilla de la capcalera: es l'unic .docx que no es una vista
    # generada. LLIC.json si que en te (vegeu _VistaLlicencia): l'usuari
    # necessita poder consultar el cataleg de Llicencia sense obrir el programa.
    $b = [System.IO.Path]::GetFileNameWithoutExtension([string]$jsonPath)
    return ([string]$b -like '0 CAPCALERA*')
}

# VERSIO del generador de vistes. Puja-la SEMPRE que canviï com es veuen les
# vistes: si no, les que ja existeixen es queden amb el format antic per sempre
# (la regla de sota nomes regenera quan el JSON es mes nou que el .docx, i just
# despres de generar-les el .docx sempre es el mes nou). En canviar de versio es
# regeneren totes una vegada.
#   6 -> LLIC.json tambe te vista (_VistaLlicencia)
#   1 -> primera versio (format propi, amb estils de titol)
#   2 -> format de l'informe (Format.ps1) + nivells d'esquema
#   3 -> tipografia base de la plantilla (Bookman Old Style, justificat,
#        interlineat i marges) via Format-ApplyBaseStyle
#   4 -> separacio entre l'item i el seu PRIMER sub-punt (Format-Bullet -First)
#   5 -> negreta del numero de l'item aplicada pel RANG (no s'encomana al cos)
#        i sangria dels fills a 1 cm amb francesa de 0,5 cm
$Script:VistaWordVersio = 7

function _VistaVersioPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return (Join-Path $base (Join-Path 'InformesCornella' 'vistes-versio.txt'))
}

# La versio amb que es van generar les vistes d'aquest ordinador (0 si no consta).
function _VistaVersioDesada {
    $p = _VistaVersioPath
    if (-not (Test-Path -LiteralPath $p)) { return 0 }
    try { return [int](Get-Content -LiteralPath $p -Raw -ErrorAction Stop).Trim() } catch { return 0 }
}

function _VistaDesaVersio([int]$v) {
    try {
        $p = _VistaVersioPath
        $d = Split-Path -Parent $p
        if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
        [string]$v | Set-Content -LiteralPath $p -Encoding UTF8
    } catch { }
}

# Cal tornar a generar la vista? Funcio PURA (mateixa forma que _PdfShouldConvert).
# NOMES es regenera si el JSON s'ha tocat despres del .docx: si es regenerava
# sempre, cada Actualitzar.bat faria un commit d'un .docx "nou" (Word hi posa
# dates internes) i el repositori s'ompliria de canvis inutils.
function _VistaCalRegenerar([bool]$docxExists, [datetime]$jsonUtc, [datetime]$docxUtc, [bool]$force) {
    if ($force) { return $true }
    if (-not $docxExists) { return $true }
    return ($jsonUtc -gt $docxUtc)
}

# D'una capcalera h2 d'ACT_EXTR ("[[CLAU]] ::TOKEN:: etiqueta") en treu una
# etiqueta llegible per al titol de la vista. Funcio PURA.
function _VistaActExtrTitol([string]$h2) {
    $t = [string]$h2
    $clau = ''
    $m = [regex]::Match($t, '^\s*\[\[([^\]]*)\]\]')
    if ($m.Success) { $clau = $m.Groups[1].Value; $t = $t.Substring($m.Index + $m.Length) }
    $t = [regex]::Replace($t, '::[A-Z]+::', '')
    $t = $t.Trim()
    if ([string]::IsNullOrWhiteSpace($t)) { return $clau }
    if ([string]::IsNullOrWhiteSpace($clau)) { return $t }
    return ($t + '  [' + $clau + ']')
}

# ----------------------------------------------------------------------------
# ESCRIPTURA A WORD (COM) - nomes Windows
# ----------------------------------------------------------------------------
# El format es EXACTAMENT el de l'informe: totes les funcions de sota criden les
# Format-* de Format.ps1 (les mateixes que fa servir Build-Document), aixi la
# vista es veu igual que sortiria el document generat.
#
# A mes, cada titol rep un NIVELL D'ESQUEMA (OutlineLevel) perque surti al panell
# de navegacio del Word. L'OutlineLevel NO canvia com es veu el paragraf: nomes
# el fa navegable. Compte: el Word HERETA el nivell al paragraf seguent, per aixo
# el cos el torna sempre a 10 (wdOutlineLevelBodyText).
# $Script:WdOutlineBody i Format-Nivell viuen a Format.ps1: tocar el Word per
# format es cosa d'aquell modul, i aixi no en queda cap copia aqui.
function _VistaNivell($sel, [int]$n) { Format-Nivell $sel $n }

# --- Embolcalls: format de l'informe + nivell d'esquema ---------------------
function _VSection($sel, [string]$t)  { Format-Section $sel $t;    _VistaNivell $sel 1 }
function _VSubsection($sel, [string]$t) { Format-Subsection $sel $t; _VistaNivell $sel 2 }
function _VItem($sel, [string]$num, [string]$t) { Format-Item $sel $num $t; _VistaNivell $sel 3 }
function _VBody($sel, [string]$t, [bool]$isChild = $false) {
    if ($isChild) { Format-Body $sel $t -IsChild } else { Format-Body $sel $t }
    _VistaNivell $sel $Script:WdOutlineBody
}
function _VUrl($sel, [string]$u, [bool]$isChild = $false) {
    if ($isChild) { Format-Url $sel $u -IsChild } else { Format-Url $sel $u }
    _VistaNivell $sel $Script:WdOutlineBody
}
function _VBullet($sel, [string]$t, [bool]$isChild = $true, [bool]$first = $false) {
    if ($isChild) { Format-Bullet $sel $t -IsChild -First:$first } else { Format-Bullet $sel $t -First:$first }
    _VistaNivell $sel $Script:WdOutlineBody
}
function _VNote($sel, [string]$t)  { Format-Note $sel $t;  _VistaNivell $sel $Script:WdOutlineBody }
function _VLabel($sel, [string]$t) { Format-Label $sel $t; _VistaNivell $sel $Script:WdOutlineBody }
function _VConcl($sel, [string]$t) { Format-Conclusion $sel $t; _VistaNivell $sel $Script:WdOutlineBody }
function _VConclCap($sel, [string]$t) { Format-ConclusionHeader $sel $t; _VistaNivell $sel $Script:WdOutlineBody }
function _VSpacer($sel) { Format-Spacer $sel; _VistaNivell $sel $Script:WdOutlineBody }
# I la versio per NOM de bloc (Format-Aire), que es la que decideix si hi va
# aire o no. Vegeu $Script:AireFlagPerClau a Format.ps1.
function _VAire($sel, [string]$clau) { if (Test-FormatAire $clau) { _VSpacer $sel } }

# Escriu una linia de cos separant text i URLs, com fa el motor (_SplitTextAndUrls).
#
# PER QUE NO ES Write-Linia (MotorInforme.ps1), que fa exactament aixo: perque
# aqui cada paragraf ha de rebre a mes el seu NIVELL D'ESQUEMA, i el Word
# l'HERETA del paragraf anterior. Posar-lo un sol cop al final deixaria els
# d'abans al nivell equivocat, o sigui que ha d'anar paragraf a paragraf
# (_VBody / _VUrl). Es fon amb la resta quan el motor de blocs sapiga de
# nivells; vegeu la fase 2 del pla.
function _VLine($sel, [string]$line, [bool]$isChild = $false) {
    if ([string]::IsNullOrWhiteSpace($line)) { return }
    $parts = _SplitTextAndUrls $line
    if (-not [string]::IsNullOrWhiteSpace($parts.Text)) { _VBody $sel $parts.Text $isChild }
    foreach ($u in $parts.Urls) { _VUrl $sel $u $isChild }
}

# ---- Vista del cataleg de LLICENCIA ----------------------------------------
# Ensenya el que Llicencia produira: cada bloc (ABANS / PROJECTE / DESPRES /
# PROPIS / ANNEX 1) amb TOTS els seus punts, el text que ve de REQ1 i, a sota,
# el que hi afegeix LLIC (els comentaris "No es disposa..." / "Es disposa..." i
# el "Quan:"). Es la manera de veure d'una ullada d'on surt cada punt.
#
# Els punts surten de _LlicPuntsPerBloc, o sigui de la MATEIXA funcio que munta
# l'informe: la vista no pot dir una cosa i el document una altra.
# La vista dels dos informes CURTS de llicencia (MNSTRAS.json). Ensenya cada un
# amb les DUES variants -amb observacions i sense-, que es l'unica cosa que hi
# canvia, i marca on va la llista que l'usuari omple al Word.
function _VistaMnsTraspas($sel, [string]$jsonPath) {
    $cfg = $Script:ReportFormatConfig
    $cat = Read-MnsCataleg $jsonPath
    foreach ($f in @(_MnsFases)) {
        _VSection $sel ([string]$f.Nom)
        _VAire $sel 'seccio'
        foreach ($v in @(@{ Amb = $false; Nom = 'sense observacions' }, @{ Amb = $true; Nom = 'amb observacions' })) {
            _VSubsection $sel ([string]$v.Nom)
            _VAire $sel 'subseccio'
            foreach ($p in @(_MnsParagrafs $cat ([string]$f.Clau) ([bool]$v.Amb))) {
                if ([string]$p.Tipus -eq 'llista') {
                    _VBody $sel ('//(aqui hi va una llista de Word buida, per omplir-la a ma)//')
                    continue
                }
                foreach ($l in @($p.Linies)) {
                    $pp = _SplitTextAndUrls ([string]$l)
                    if (-not [string]::IsNullOrWhiteSpace($pp.Text)) { _VBody $sel $pp.Text }
                }
            }
            _VAire $sel 'item'
        }
    }
}

function _VistaLlicencia($sel, [string]$jsonPath) {
    $cfg = $Script:ReportFormatConfig
    $llic = Read-LlicCataleg $jsonPath
    $req1Path = Join-Path (Split-Path -Parent $jsonPath) 'REQ1.json'
    $req1 = if (Test-Path -LiteralPath $req1Path) { Read-CatalegJson $req1Path } else { $null }
    $idx = _LlicIndexReq1 $req1

    $blocs = @(
        @{ Clau = 'PROPIS';  Titol = 'PUNTS PROPIS DE LLIC' + [char]0x00C8 + 'NCIA (no son a REQ1)' },
        @{ Clau = 'ABANS';   Titol = (_LlicTitolAbans) },
        @{ Clau = 'DESPRES'; Titol = (_LlicTitolDespres) }
    )
    foreach ($b in $blocs) {
        $r = _LlicPuntsPerBloc $llic $idx ([string]$b.Clau) $req1
        _VSection $sel ([string]$b.Titol)
        _VAire $sel 'seccio'
        $seccioAra = ''
        $n = 0
        foreach ($p in @($r.Punts)) {
            $sec = _LlicSeccioDePunt $p
            if ($sec -ne $seccioAra) {
                $seccioAra = $sec
                if ($sec) {
                    _VSubsection $sel ('de REQ1: ' + $sec)
                    _VAire $sel 'subseccio'
                }
            }
            $linies = @($p.Cos)
            $n++
            if ($linies.Count -gt 0) {
                $p0 = _SplitTextAndUrls ([string]$linies[0])
                _VItem $sel ("$n.") ([string]$p0.Text)
                foreach ($u in $p0.Urls) { _VUrl $sel $u }
                for ($i = 1; $i -lt $linies.Count; $i++) { _VLine $sel ([string]$linies[$i]) }
            } else {
                _VItem $sel ("$n.") ([string]$p.Titol)
            }
            foreach ($sub in @($p.Subs)) {
                foreach ($l in @($sub)) { _VLine $sel ([string]$l) $true }
            }
            foreach ($par in @(
                @{ E = 'No es disposa'; L = @($p.NoDisposa) },
                @{ E = 'Es disposa';    L = @($p.SiDisposa) },
                @{ E = 'Quan';          L = @($p.Quan) })) {
                foreach ($l in @($par.L)) {
                    if ([string]::IsNullOrWhiteSpace([string]$l)) { continue }
                    _VLine $sel ('//[' + [string]$par.E + ']// ' + [string]$l) $true
                }
            }
            _VAire $sel 'item'
        }
        if (@($r.Orfes).Count -gt 0) {
            _VBody $sel ('**Claus que ja NO son a REQ1: ' + (@($r.Orfes) -join ' | ') + '**')
            _VAire $sel 'item'
        }
    }

    # El PROJECTE: la resta de REQ1, la que no es demana ni abans ni despres.
    if ($null -ne $req1) {
        _VSection $sel 'PROJECTE (la resta de REQ1)'
        _VAire $sel 'seccio'
        $senseAbans = @(@($req1.Sections) | Where-Object { -not (_LlicEsSeccioAbans ([string]$_.Title)) })
        $secProj = @(_LlicSeccionsSenseSubseccions $senseAbans (_LlicSeccionsExpandides $llic $idx))
        foreach ($sc in $secProj) {
            _VBody $sel ('//' + [string]$sc.Title + ' (' +
                         @($sc.Items | Where-Object { [string]$_.Kind -eq 'item' }).Count + ' punts)//')
        }
        _VAire $sel 'item'
    }

    # L'ANNEX 1, tal com surt a l'informe.
    $secAnnex = _LlicSeccioAnnex1 $llic
    if ($null -ne $secAnnex) {
        _VSection $sel ([string]$secAnnex.titol)
        _VAire $sel 'seccio'
        $num = 0
        foreach ($nd in @($secAnnex.fills)) {
            $marca = ''
            $tip = [string]$nd.tipus
            if ($tip -eq 'item') { $num++; $marca = [string]$num + '. ' }
            elseif ($tip -eq 'subitem') { $marca = '- ' }
            $primera = $true
            foreach ($l in @(_LlicCos $nd)) {
                _VLine $sel ($(if ($primera) { $marca } else { '' }) + [string]$l)
                $primera = $false
            }
        }
    }
}

# ---- Vista d'un CATALEG (REQ1, TERMINI...) ---------------------------------
# Reprodueix el que faria _WriteCatalegBody amb TOTS els items triats: seccio en
# MAJUSCULES, subseccio subratllada, items numerats amb el numero en negreta i
# fills com a punts de llista. Els [CAMP:]/[OPCIO:] es deixen tal qual (es una
# vista del cataleg, no un informe d'una activitat concreta).
function _VistaCataleg($sel, [string]$jsonPath, [string]$nom) {
    $parsed = Read-CatalegJson $jsonPath
    $cfg = $Script:ReportFormatConfig

    if (-not [string]::IsNullOrWhiteSpace([string]$parsed.IntroText)) {
        _VBody $sel ([string]$parsed.IntroText)
        _VAire $sel 'introparagraf'
    }
    if ($parsed.IsFixedBody) {
        $lines = @($parsed.FixedBodyLines)
        for ($i = 0; $i -lt $lines.Count; $i++) {
            _VLine $sel ([string]$lines[$i])
            if ($i -lt ($lines.Count - 1)) { _VSpacer $sel }
        }
        return
    }

    $num = 0
    foreach ($sec in @($parsed.Sections)) {
        _VSection $sel ([string]$sec.Title)
        _VAire $sel 'seccio'
        foreach ($el in @($sec.Items)) {
            switch ([string]$el.Kind) {
                'subsection' {
                    _VSubsection $sel ([string]$el.Short)
                    _VAire $sel 'subseccio'
                }
                'intro' {
                    foreach ($ln in @($el.BodyLines)) { _VLine $sel ([string]$ln) }
                    _VAire $sel 'intro'
                }
                default {
                    $lines = @($el.BodyLines)
                    $escrit = $false
                    if ($lines.Count -gt 0) {
                        $num++
                        $p0 = _SplitTextAndUrls ([string]$lines[0])
                        _VItem $sel ("$num.") ([string]$p0.Text)
                        foreach ($u in $p0.Urls) { _VUrl $sel $u }
                        for ($i = 1; $i -lt $lines.Count; $i++) { _VLine $sel ([string]$lines[$i]) }
                        $escrit = $true
                    }
                    $primerFill = $true
                    foreach ($ch in @($el.Children)) {
                        $cl = @($ch.BodyLines)
                        if ($cl.Count -eq 0) { continue }
                        if (-not $escrit) { $num++; $escrit = $true }
                        # Els fills NO es numeren: van amb pic, com a l'informe.
                        # El PRIMER se separa mes de l'item (com fa Motor.ps1).
                        $pc = _SplitTextAndUrls ([string]$cl[0])
                        if (-not [string]::IsNullOrWhiteSpace($pc.Text)) {
                            _VBullet $sel ([string]$pc.Text) $true $primerFill
                            $primerFill = $false
                        }
                        foreach ($u in $pc.Urls) { _VUrl $sel $u $true }
                        for ($i = 1; $i -lt $cl.Count; $i++) { _VLine $sel ([string]$cl[$i]) $true }
                    }
                    if ($escrit) { _VAire $sel 'item' }
                }
            }
        }
    }
}

# ---- Vista de les CONCLUSIONS ----------------------------------------------
function _VistaConclusions($sel, [string]$jsonPath) {
    $o = _LoadEstructuralJson $jsonPath
    $sempre = New-Object System.Collections.ArrayList
    foreach ($p in @($o.intro)) {
        $t = _JsonParaToBodyLine $p
        if (-not [string]::IsNullOrWhiteSpace($t)) { _VSection $sel $t }
    }
    foreach ($n in @($o.nodes)) {
        if ([string]$n.tipus -eq 'sempre') {
            foreach ($p in @($n.cos)) { [void]$sempre.Add((_JsonParaToBodyLine $p)) }
            continue
        }
        # Grup = tipus d'informe (REQ1, SEGUIMENT, TERMINI...).
        _VSpacer $sel
        _VSection $sel ('Conclusions de ' + [string]$n.titol)
        _VSpacer $sel
        $num = 0
        foreach ($c in @($n.fills)) {
            $num++
            $cos = @($c.cos)
            $primera = if ($cos.Count -gt 0) { _JsonParaToBodyLine $cos[0] } else { '' }
            _VItem $sel ("$num.") ([string]$primera)
            for ($i = 1; $i -lt $cos.Count; $i++) { _VLine $sel (_JsonParaToBodyLine $cos[$i]) }
            _VSpacer $sel
        }
    }
    if ($sempre.Count -gt 0) {
        _VSpacer $sel
        _VSection $sel 'Frases que surten sempre'
        _VSpacer $sel
        foreach ($l in $sempre) { _VLine $sel ([string]$l) }
    }
}

# ---- Vista d'una plantilla ACT_EXTR ----------------------------------------
# EL CONTINGUT ES PINTA COM AL DOCUMENT, segons el TOKEN del bloc
# (::CHILD::, ::NOTE::, ::LABEL::, ::HEADER::, ::CONC::, ::TEXT:: o res = item),
# el mateix que fa _WriteActExtrBody.
#
# Abans la vista ho numerava TOT -tambe els sub-punts, les notes, les etiquetes
# i les conclusions de l'informe favorable-, o sigui que ensenyava una cosa i el
# document en generava una altra. Una vista que no s'assembla al que surt no
# serveix per consultar-la, que es tot el motiu de tenir-la.
function _VActExtrContingut($sel, [string]$kind, [string]$txt, [ref]$primerFill) {
    if ([string]::IsNullOrWhiteSpace($txt)) { return }
    if ($kind -eq 'child') {
        _VBullet $sel $txt $true ([bool]$primerFill.Value)
        $primerFill.Value = $false
        return
    }
    switch ($kind) {
        'note'   { _VNote $sel $txt }
        'label'  { _VLabel $sel $txt }
        'header' { _VConclCap $sel $txt }
        'conc'   { _VConcl $sel $txt }
        'text'   { _VBody $sel $txt }
        default  { _VBullet $sel $txt $false }   # 'item': pic de primer nivell
    }
    $primerFill.Value = $true
}

function _VistaActExtr($sel, [string]$jsonPath, [string]$nom) {
    $records = @(Read-ActExtrRecordsJson $jsonPath)
    $kind = 'item'
    $primerFill = $false
    foreach ($r in $records) {
        $txt = [string]$r.Text
        switch ([string]$r.Style) {
            'h1' {
                _VSpacer $sel
                _VSection $sel (_VistaActExtrTitol $txt)
                _VSpacer $sel
                $mk = _ParseActExtrMarker $txt
                $kind = if ($null -ne $mk) { [string]$mk.Kind } else { 'item' }
            }
            'h2' {
                # La capcalera del bloc ("[[CLAU]] ::TOKEN:: etiqueta") no surt a
                # l'informe: al document nomes hi va el CONTINGUT. A la vista si
                # que la posem (subratllada) per saber quin bloc es cadascun.
                _VSubsection $sel (_VistaActExtrTitol $txt)
                $mk = _ParseActExtrMarker $txt
                $kind = if ($null -ne $mk) { [string]$mk.Kind } else { 'item' }
                # $primerFill NO es reinicia aqui: al document les capcaleres de
                # bloc no escriuen res, o sigui que el primer sub-punt d'un bloc
                # 'child' segueix penjant de la unitat anterior. Reiniciar-lo
                # faria que a la vista cap sub-punt no sortis mai com a primer.
            }
            'url' { _VUrl $sel $txt ($kind -eq 'child') }
            default {
                $pf = $primerFill
                _VActExtrContingut $sel $kind $txt ([ref]$pf)
                $primerFill = $pf
            }
        }
    }
}

# ---- Genera la vista d'UN cataleg ------------------------------------------
function Export-VistaWord($word, [string]$jsonPath) {
    if (_VistaEsProtegit $jsonPath) { return $false }
    if (-not (Test-Path -LiteralPath $jsonPath)) { return $false }
    $o = _LoadEstructuralJson $jsonPath
    $familia = [string]$o.familia
    $nom = [System.IO.Path]::GetFileNameWithoutExtension($jsonPath)
    # [string] EXPLICIT: SaveAs rep la ruta per REFERENCIA ([ref]$out) i, si el
    # valor ve embolcallat en un PSObject, el COM no el sap convertir.
    [string]$out = _VistaWordPathFor $jsonPath
    # La carpeta de les vistes pot no existir encara (clone acabat de baixar).
    $outDir = Split-Path -Parent $out
    if (-not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $doc = $word.Documents.Add()
    try {
        # Un document NOU de Word surt en Calibri, alineat a l'esquerra i amb
        # uns altres marges. Li posem la MATEIXA base que la plantilla de
        # l'informe (Bookman Old Style, justificat, interlineat i marges), que
        # es on esta declarada: Format.ps1.
        Format-ApplyBaseStyle $doc
        $sel = $word.Selection
        switch ($familia) {
            'cataleg'     { _VistaCataleg $sel $jsonPath $nom }
            'conclusions' { _VistaConclusions $sel $jsonPath }
            'actextr'     { _VistaActExtr $sel $jsonPath $nom }
            'llicencia'   { _VistaLlicencia $sel $jsonPath }
            'mnstraspas'  { _VistaMnsTraspas $sel $jsonPath }
            default       { _VistaCataleg $sel $jsonPath $nom }
        }
        # Nota final: que quedi clar que es una vista generada i que no s'edita.
        _VSpacer $sel
        _VBody $sel ("//Vista generada autom" + [char]0x00E0 + "ticament des de " + [System.IO.Path]::GetFileName($jsonPath) + " el " + (Get-Date).ToString('dd/MM/yyyy HH:mm') + ". No l'editis: els canvis es fan des de l'editor de cat" + [char]0x00E0 + "legs del programa.//")
        $doc.SaveAs([ref]$out, [ref]16)   # 16 = wdFormatDocumentDefault (.docx)
        return $true
    } finally {
        try { $doc.Close($false) } catch { }
    }
}

# ---- Genera TOTES les vistes ------------------------------------------------
# Retorna el nombre de vistes generades. Fail-safe: si no hi ha Word, no peta.
function Invoke-ExportarVistesWord([switch]$Force) {
    if (-not (Test-Path -LiteralPath $EstructuralsDir)) { return 0 }
    $tots = @(Get-ChildItem -LiteralPath $EstructuralsDir -Filter '*.json' -ErrorAction SilentlyContinue |
              Where-Object { -not (_VistaEsProtegit $_.FullName) } | Sort-Object Name)
    # Si ha canviat el format de les vistes, es regeneren TOTES una vegada.
    $canviDeFormat = ((_VistaVersioDesada) -ne $Script:VistaWordVersio)
    $forcar = ([bool]$Force -or $canviDeFormat)
    if ($canviDeFormat) { Write-Host "  (el format de les vistes ha canviat: es regeneren totes)" }

    # Nomes els que tinguin el JSON mes nou que la vista (o cap vista encara).
    $jsons = @()
    foreach ($j in $tots) {
        $out = _VistaWordPathFor $j.FullName
        $ex = Test-Path -LiteralPath $out
        $docxUtc = if ($ex) { (Get-Item -LiteralPath $out).LastWriteTimeUtc } else { [datetime]::MinValue }
        if (_VistaCalRegenerar $ex $j.LastWriteTimeUtc $docxUtc $forcar) { $jsons += $j }
    }
    if ($jsons.Count -eq 0) {
        if ($canviDeFormat) { _VistaDesaVersio $Script:VistaWordVersio }
        return 0
    }

    $word = $null
    try { $word = New-Object -ComObject Word.Application } catch { $word = $null }
    if ($null -eq $word) {
        Write-Host "  Avis: no s'ha pogut obrir el Word; no s'han generat les vistes."
        return 0
    }
    $word.Visible = $false
    try { $word.DisplayAlerts = 0 } catch { }
    $n = 0
    try {
        foreach ($j in $jsons) {
            try {
                if (Export-VistaWord $word $j.FullName) {
                    $n++
                    Write-Host ("  vista: " + [System.IO.Path]::GetFileNameWithoutExtension($j.Name) + ".docx")
                }
            } catch {
                Write-Host ("  Avis: no s'ha pogut generar la vista de " + $j.Name + " (" + $_.Exception.Message + ")")
            }
        }
    } finally {
        try { $word.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    }
    # Nomes donem la versio per bona si s'han pogut generar (si Word ha fallat,
    # la propera vegada ho tornara a intentar).
    if ($n -gt 0) { _VistaDesaVersio $Script:VistaWordVersio }
    return $n
}
