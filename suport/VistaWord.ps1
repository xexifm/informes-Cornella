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
#   9 -> la FITXA D'AJUDA de cada requeriment, en gris (Format-Ajuda)
#  10 -> la fitxa sense sangria (com la resta del cos), amb la VIGENCIA i amb
#        l'ENLLAC al text consolidat de la norma com a hipervincle
#  11 -> la de LLIC en blocs (els enllacos surten com a enllacos)
#  12 -> ACT_EXTR, MNS i conclusions tambe en blocs; amb l'aire apagat, els
#        titols ja no perden el nivell d'esquema (sortien del panell)
$Script:VistaWordVersio = 12

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
# TOTES LES VISTES SON BLOCS + Write-Informe -AmbNivells
# ----------------------------------------------------------------------------
# El format es EXACTAMENT el de l'informe perque qui escriu es el mateix motor
# (Write-Informe, MotorInforme.ps1). -AmbNivells hi afegeix el NIVELL D'ESQUEMA
# (OutlineLevel) de cada paragraf perque la vista sigui navegable des del panell
# del Word; no en canvia l'aspecte.
#
# Abans les vistes d'ACT_EXTR, MNS/Traspas i conclusions escrivien amb onze
# embolcalls propis (_VSection, _VBody, _VBullet...) que repetien, un per un, el
# que ja fa el motor. Ara cada vista es un Build-*VistaBlocs PUR (es prova a
# Linux comptant blocs) i la funcio _Vista* nomes llegeix el cataleg i escriu.

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
function Build-MnsVistaBlocs($cat) {
    $b = New-Object System.Collections.ArrayList
    foreach ($f in @(_MnsFases)) {
        [void]$b.Add(@{ T = 'seccio'; Text = [string]$f.Nom })
        [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
        foreach ($v in @(@{ Amb = $false; Nom = 'sense observacions' }, @{ Amb = $true; Nom = 'amb observacions' })) {
            [void]$b.Add(@{ T = 'subseccio'; Text = [string]$v.Nom })
            [void]$b.Add(@{ T = 'aire'; Clau = 'subseccio' })
            foreach ($p in @(_MnsParagrafs $cat ([string]$f.Clau) ([bool]$v.Amb))) {
                if ([string]$p.Tipus -eq 'llista') {
                    [void]$b.Add(@{ T = 'cos'; Text = '//(aqui hi va una llista de Word buida, per omplir-la a ma)//' })
                    continue
                }
                # Nomes el TEXT: la vista d'aquests dos informes mai no ha
                # ensenyat els enllacos (no en porten).
                foreach ($l in @($p.Linies)) {
                    $pp = _SplitTextAndUrls ([string]$l)
                    if (-not [string]::IsNullOrWhiteSpace($pp.Text)) { [void]$b.Add(@{ T = 'cos'; Text = [string]$pp.Text }) }
                }
            }
            [void]$b.Add(@{ T = 'aire'; Clau = 'item' })
        }
    }
    return $b.ToArray()
}

function _VistaMnsTraspas($sel, [string]$jsonPath) {
    [void](Write-Informe $sel (Build-MnsVistaBlocs (Read-MnsCataleg $jsonPath)) -AmbNivells)
}

# UN PUNT a la VISTA de LLIC: el text de REQ1 (o el propi) i, a sota i en
# cursiva, el que hi afegeix LLIC ([No es disposa], [Es disposa], [Quan]). Els
# [CAMP:] es veuen tal qual. Funcio PURA.
#
# EL COMENTARI ES PARTEIX EN TEXT I ENLLAC ABANS DE POSAR-HI L'ETIQUETA. Abans
# s'hi enganxava "//[No es disposa]// " al davant de la linia sencera, i una
# linia que era NOMES un enllac ("[[URL]] https://...") deixava de comencar per
# [[URL]]: a la vista sortia "[No es disposa] [[URL]]" escrit com a text. Es veu
# al fitxer d'or de la vista d'abans d'unificar-la amb l'informe.
function _LlicVistaBlocsDePunt($p, [string]$marca) {
    $out = New-Object System.Collections.ArrayList
    $linies = @($p.Cos)
    if ($linies.Count -gt 0) {
        $p0 = _SplitTextAndUrls ([string]$linies[0])
        [void]$out.Add(@{ T = 'item'; Num = $marca; Text = [string]$p0.Text })
        foreach ($u in @($p0.Urls)) { [void]$out.Add(@{ T = 'enllac'; Url = $u }) }
        for ($i = 1; $i -lt $linies.Count; $i++) { foreach ($x in @(_BlocsDeLinia ([string]$linies[$i]) $false)) { [void]$out.Add($x) } }
    } else {
        [void]$out.Add(@{ T = 'item'; Num = $marca; Text = [string]$p.Titol })
    }
    foreach ($sub in @($p.Subs)) {
        foreach ($l in @($sub)) { foreach ($x in @(_BlocsDeLinia ([string]$l) $true)) { [void]$out.Add($x) } }
    }
    foreach ($par in @(
        @{ E = 'No es disposa'; L = @($p.NoDisposa) },
        @{ E = 'Es disposa';    L = @($p.SiDisposa) },
        @{ E = 'Quan';          L = @($p.Quan) })) {
        foreach ($l in @($par.L)) {
            if ([string]::IsNullOrWhiteSpace([string]$l)) { continue }
            $pp = _SplitTextAndUrls ([string]$l)
            if (-not [string]::IsNullOrWhiteSpace($pp.Text)) {
                [void]$out.Add(@{ T = 'cos'; Text = ('//[' + [string]$par.E + ']// ' + [string]$pp.Text); Fill = $true })
            }
            foreach ($u in @($pp.Urls)) { [void]$out.Add(@{ T = 'enllac'; Url = $u; Fill = $true }) }
        }
    }
    return $out.ToArray()
}

# LA VISTA DE LLIC.json, en blocs. Funcio PURA.
#
# Cada bloc (PROPIS / ABANS / DESPRES) amb TOTS els seus punts, amb la MATEIXA
# estructura que l'informe: la fa _LlicBlocsPunts, la mateixa funcio. Abans la
# vista en tenia una copia i s'havia de tocar alhora que l'informe i que REQ1.
function Build-LlicVistaBlocs($llic, $req1) {
    $b = New-Object System.Collections.ArrayList
    $idx = _LlicIndexReq1 $req1
    $puntVista = { param($p, $marca) _LlicVistaBlocsDePunt $p $marca }
    foreach ($bl in @(
        @{ Clau = 'PROPIS';  Titol = 'PUNTS PROPIS DE LLIC' + [char]0x00C8 + 'NCIA (no son a REQ1)' },
        @{ Clau = 'ABANS';   Titol = (_LlicTitolAbans) },
        @{ Clau = 'DESPRES'; Titol = (_LlicTitolDespres) })) {
        $r = _LlicPuntsPerBloc $llic $idx ([string]$bl.Clau) $req1
        [void]$b.Add(@{ T = 'titolbloc'; Text = [string]$bl.Titol })
        [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
        $n = 0   # a la vista, cada bloc numera des de l'1
        foreach ($x in @(_LlicBlocsPunts @($r.Punts) ([ref]$n) $null 'numero' $puntVista $true)) { [void]$b.Add($x) }
        if (@($r.Orfes).Count -gt 0) {
            $txt = '**Claus que ja NO son a REQ1: ' + (@($r.Orfes) -join ' | ') + '**'
            [void]$b.Add(@{ T = 'unitat'; Blocs = @(@{ T = 'cos'; Text = $txt }) })
        }
    }

    # El PROJECTE: la resta de REQ1, la que no es demana ni abans ni despres.
    if ($null -ne $req1) {
        [void]$b.Add(@{ T = 'seccio'; Text = 'PROJECTE (la resta de REQ1)' })
        [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
        $senseAbans = @(@($req1.Sections) | Where-Object { -not (_LlicEsSeccioAbans ([string]$_.Title)) })
        $secProj = @(_LlicSeccionsSenseSubseccions $senseAbans (_LlicSeccionsExpandides $llic $idx))
        $u = New-Object System.Collections.ArrayList
        foreach ($sc in $secProj) {
            [void]$u.Add(@{ T = 'cos'; Text = ('//' + [string]$sc.Title + ' (' + @($sc.Items | Where-Object { [string]$_.Kind -eq 'item' }).Count + ' punts)//') })
        }
        [void]$b.Add(@{ T = 'unitat'; Blocs = $u.ToArray() })
    }

    # QUI POSA CONDICIONS, amb el punt de REQ1 que en proposa cada un. Surt de
    # _LlicActorsCondicions, la mateixa funcio que la pantalla.
    $actorsV = @(_LlicActorsCondicions $llic)
    if ($actorsV.Count -gt 0) {
        [void]$b.Add(@{ T = 'seccio'; Text = 'CONDICIONS (qui les posa)' })
        [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
        $ia = 0
        foreach ($a in $actorsV) {
            $ia++
            $u = New-Object System.Collections.ArrayList
            [void]$u.Add(@{ T = 'item'; Num = ((_LlicLletra $ia).ToLower() + '.'); Text = [string]$a.Nom })
            foreach ($c in @($a.Claus)) { [void]$u.Add(@{ T = 'cos'; Text = ('//[Es proposa si es disposa de]// ' + [string]$c); Fill = $true }) }
            [void]$b.Add(@{ T = 'unitat'; Blocs = $u.ToArray() })
        }
    }

    # L'ANNEX 1 (a la vista, text corrent amb la marca al davant).
    $secAnnex = _LlicSeccioAnnex1 $llic
    if ($null -ne $secAnnex) {
        [void]$b.Add(@{ T = 'seccio'; Text = [string]$secAnnex.titol })
        [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
        $num = 0
        foreach ($nd in @($secAnnex.fills)) {
            $marca = ''
            $tip = [string]$nd.tipus
            if ($tip -eq 'item') { $num++; $marca = [string]$num + '. ' }
            elseif ($tip -eq 'subitem') { $marca = '- ' }
            $primera = $true
            foreach ($l in @(_LlicCos $nd)) {
                foreach ($x in @(_BlocsDeLinia ($(if ($primera) { $marca } else { '' }) + [string]$l) $false)) { [void]$b.Add($x) }
                $primera = $false
            }
        }
    }
    return $b.ToArray()
}

function _VistaLlicencia($sel, [string]$jsonPath) {
    $llic = Read-LlicCataleg $jsonPath
    $req1Path = Join-Path (Split-Path -Parent $jsonPath) 'REQ1.json'
    $req1 = if (Test-Path -LiteralPath $req1Path) { Read-CatalegJson $req1Path } else { $null }
    [void](Write-Informe $sel (Build-LlicVistaBlocs $llic $req1) -AmbNivells)
}

# ---- Vista d'un CATALEG (REQ1, TERMINI...) ---------------------------------
# Reprodueix el que faria _WriteCatalegBody amb TOTS els items triats: seccio en
# MAJUSCULES, subseccio subratllada, items numerats amb el numero en negreta i
# fills com a punts de llista. Els [CAMP:]/[OPCIO:] es deixen tal qual (es una
# vista del cataleg, no un informe d'una activitat concreta).
function _VistaCataleg($sel, [string]$jsonPath, [string]$nom) {
    $parsed = Read-CatalegJson $jsonPath

    # LA MATEIXA FUNCIO QUE L'INFORME. Abans aqui hi havia una copia de
    # _WriteCatalegBody -seccio, subseccio, items numerats, fills amb pic...- i
    # una copia podia dir una cosa mentre el document en generava una altra.
    # Ara nomes canvia una cosa: -AmbNivells, que posa l'OutlineLevel a cada
    # paragraf perque la vista sigui navegable des del panell del Word.
    #
    # A la vista hi SURT TOT: els items del cataleg no venen d'una tria, o sigui
    # que se'ls marca com a triats perque el motor els escrigui tots.
    foreach ($sec in @($parsed.Sections)) {
        foreach ($el in @($sec.Items)) {
            if ([string]$el.Kind -eq 'subsection' -or [string]$el.Kind -eq 'intro') { continue }
            $el | Add-Member NoteProperty Selected $true -Force
        }
    }
    # -SenseCamps: a la vista els [CAMP:]/[OPCIO:] es veuen TAL QUAL. Es una
    # vista del CATALEG, no l'informe d'una activitat: resoldre'ls amb un
    # diccionari buit els deixaria en blanc i la vista perdria el que hi vas a
    # mirar.
    # -AmbAjuda: la vista es el document que l'inspector consulta quan dubta si
    # ha de requerir una cosa o no, o sigui que es EL LLOC de la fitxa de
    # criteri. A l'informe del titular no hi arriba: aquest interruptor nomes el
    # posa aqui.
    $blocs = Build-CatalegBlocs $parsed.Sections $null ([string]$parsed.IntroText) ([bool]$parsed.IsFixedBody) @($parsed.FixedBodyLines) -SenseCamps -AmbAjuda
    [void](Write-Informe $sel $blocs -AmbNivells)
}

# ---- Vista de les CONCLUSIONS ----------------------------------------------
# $o es el JSON de 0 CONCLUSIONS.json tal com el torna _LoadEstructuralJson.
function Build-ConclusionsVistaBlocs($o) {
    $b = New-Object System.Collections.ArrayList
    $sempre = New-Object System.Collections.ArrayList
    foreach ($p in @($o.intro)) {
        $t = _JsonParaToBodyLine $p
        if (-not [string]::IsNullOrWhiteSpace($t)) { [void]$b.Add(@{ T = 'seccio'; Text = [string]$t }) }
    }
    foreach ($n in @($o.nodes)) {
        if ([string]$n.tipus -eq 'sempre') {
            foreach ($p in @($n.cos)) { [void]$sempre.Add((_JsonParaToBodyLine $p)) }
            continue
        }
        # Grup = tipus d'informe (REQ1, SEGUIMENT, TERMINI...).
        [void]$b.Add(@{ T = 'espai' })
        [void]$b.Add(@{ T = 'seccio'; Text = ('Conclusions de ' + [string]$n.titol) })
        [void]$b.Add(@{ T = 'espai' })
        $num = 0
        foreach ($c in @($n.fills)) {
            $num++
            $cos = @($c.cos)
            $primera = if ($cos.Count -gt 0) { _JsonParaToBodyLine $cos[0] } else { '' }
            [void]$b.Add(@{ T = 'item'; Num = "$num."; Text = [string]$primera })
            for ($i = 1; $i -lt $cos.Count; $i++) {
                foreach ($x in @(_BlocsDeLinia (_JsonParaToBodyLine $cos[$i]) $false)) { [void]$b.Add($x) }
            }
            [void]$b.Add(@{ T = 'espai' })
        }
    }
    if ($sempre.Count -gt 0) {
        [void]$b.Add(@{ T = 'espai' })
        [void]$b.Add(@{ T = 'seccio'; Text = 'Frases que surten sempre' })
        [void]$b.Add(@{ T = 'espai' })
        foreach ($l in $sempre) { foreach ($x in @(_BlocsDeLinia ([string]$l) $false)) { [void]$b.Add($x) } }
    }
    return $b.ToArray()
}

function _VistaConclusions($sel, [string]$jsonPath) {
    [void](Write-Informe $sel (Build-ConclusionsVistaBlocs (_LoadEstructuralJson $jsonPath)) -AmbNivells)
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
#
# Els pics porten el 'First' JA DECIDIT (l'excepcio que Write-Informe accepta
# nomes per a ACT_EXTR): el primer sub-punt es el que segueix qualsevol cosa que
# no sigui un sub-punt, i els pics de 1r nivell no el porten mai.
$Script:VistaActExtrTipus = @{
    'note' = 'nota'; 'label' = 'etiqueta'; 'header' = 'conclusiocap'; 'conc' = 'conclusio'; 'text' = 'cos'
}

# $records surt de Read-ActExtrRecordsJson. Funcio PURA.
function Build-ActExtrVistaBlocs($records) {
    $b = New-Object System.Collections.ArrayList
    $kind = 'item'
    $primerFill = $false
    foreach ($r in @($records)) {
        $txt = [string]$r.Text
        switch ([string]$r.Style) {
            'h1' {
                [void]$b.Add(@{ T = 'espai' })
                [void]$b.Add(@{ T = 'seccio'; Text = (_VistaActExtrTitol $txt) })
                [void]$b.Add(@{ T = 'espai' })
                $mk = _ParseActExtrMarker $txt
                $kind = if ($null -ne $mk) { [string]$mk.Kind } else { 'item' }
            }
            'h2' {
                # La capcalera del bloc ("[[CLAU]] ::TOKEN:: etiqueta") no surt a
                # l'informe: al document nomes hi va el CONTINGUT. A la vista si
                # que la posem (subratllada) per saber quin bloc es cadascun.
                [void]$b.Add(@{ T = 'subseccio'; Text = (_VistaActExtrTitol $txt) })
                $mk = _ParseActExtrMarker $txt
                $kind = if ($null -ne $mk) { [string]$mk.Kind } else { 'item' }
                # $primerFill NO es reinicia aqui: al document les capcaleres de
                # bloc no escriuen res, o sigui que el primer sub-punt d'un bloc
                # 'child' segueix penjant de la unitat anterior. Reiniciar-lo
                # faria que a la vista cap sub-punt no sortis mai com a primer.
            }
            'url' { [void]$b.Add(@{ T = 'enllac'; Url = $txt; Fill = ($kind -eq 'child') }) }
            default {
                # Sense break ni continue: dins d'un switch no fan el que sembla.
                if ([string]::IsNullOrWhiteSpace($txt)) {
                } elseif ($kind -eq 'child') {
                    [void]$b.Add(@{ T = 'pic'; Text = $txt; Fill = $true; First = $primerFill })
                    $primerFill = $false
                } else {
                    $tipus = $Script:VistaActExtrTipus[$kind]
                    if ($null -ne $tipus) { [void]$b.Add(@{ T = $tipus; Text = $txt }) }
                    else { [void]$b.Add(@{ T = 'pic'; Text = $txt; Fill = $false; First = $false }) }   # 'item': pic de primer nivell
                    $primerFill = $true
                }
            }
        }
    }
    return $b.ToArray()
}

function _VistaActExtr($sel, [string]$jsonPath, [string]$nom) {
    [void](Write-Informe $sel (Build-ActExtrVistaBlocs @(Read-ActExtrRecordsJson $jsonPath)) -AmbNivells)
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
        $nota = "//Vista generada autom" + [char]0x00E0 + "ticament des de " + [System.IO.Path]::GetFileName($jsonPath) + " el " + (Get-Date).ToString('dd/MM/yyyy HH:mm') + ". No l'editis: els canvis es fan des de l'editor de cat" + [char]0x00E0 + "legs del programa.//"
        [void](Write-Informe $sel @(@{ T = 'espai' }, @{ T = 'cos'; Text = $nota }) -AmbNivells)
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

    # -Opcional: aqui es pot continuar sense Word (no es genera cap vista i es
    # diu). El que s'hi guanya respecte del New-Object a pel es
    # l'AutomationSecurity de New-WordApp: sense ell, el Word obre en VISTA
    # PROTEGIDA el que ve d'una unitat de xarxa i les modificacions no tiren.
    $word = New-WordApp -Opcional
    if ($null -eq $word) {
        Write-Host "  Avis: no s'ha pogut obrir el Word; no s'han generat les vistes."
        return 0
    }
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
