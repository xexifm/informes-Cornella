#requires -Version 5.1
<#
.SYNOPSIS
  El motor de composicio d'informes: el que TOTS els informes fan igual.

.DESCRIPTION
  Aqui hi va el que abans estava escrit una vegada a cada familia d'informe.
  Format.ps1 diu COM es veu cada paragraf; aquest modul diu COM ES MUNTA un
  document sencer, de manera que canviar-ho en un lloc afecti a tots.

  Punts d'entrada:
    Write-InformeDocx  -> obrir la plantilla, escriure el cos i desar
#>

# ----------------------------------------------------------------------------
# OBRIR, ESCRIURE I DESAR UN INFORME
# ----------------------------------------------------------------------------
# La seqüencia que feien IGUAL les quatre families (REQ1/TERMINI, ACT_EXTR,
# Llicencia i MNS/Traspas), ~20 linies copiades quatre vegades:
#
#   nom unic al directori de sortida -> copia a %TEMP% (si no, el Word obre el
#   fitxer en "Vista protegida" quan el desti es una unitat de xarxa) -> triar
#   el bloc de capcalera -> substituir els <<PLACEHOLDERS>> -> escriure el cos
#   -> desar, tancar i moure al desti.
#
#   $baseName : nom de fitxer ja calculat per la familia (cada una te el seu
#               patro; l'unic que comparteixen es la data al davant).
#   $capBloc  : quin bloc de '0 CAPCALERA.docx' (''=el generic, 'ACT_EXTR',
#               'LLIC'). Si el bloc no hi es, Select-CapcaleraBlock es queda amb
#               el generic i l'informe surt igualment.
#   $cos      : scriptblock que rep la Selection ja col·locada al final del
#               document i hi escriu el cos.
#
# COMPTE: $cos NO ha de portar .GetNewClosure(). Ha de veure els locals del
# Build-* que el crea en TEMPS D'EXECUCIO, i .GetNewClosure() en copiaria els
# VALORS del moment de crear-lo (vegeu CLAUDE.md, "les dues cares de la
# closure").
#
# Si el cos peta, el document es TANCA abans de rellancar l'error: si no, es
# queda una instancia de Word amb un document obert i el %TEMP% brut. Nomes
# ACT_EXTR ho feia; les altres tres no.
function Write-InformeDocx($word, [string]$baseName, [string]$capBloc, $header, [scriptblock]$cos) {
    $targetDir = _ResolveOutputDir
    [string]$outPath = _GetUniqueOutputPath $targetDir $baseName
    $fileName = [System.IO.Path]::GetFileName($outPath)
    $tempPath = Join-Path $env:TEMP $fileName
    $doc = _OpenOutputDocument $word $tempPath
    try {
        Select-CapcaleraBlock $doc $capBloc
        Apply-HeaderReplacements -doc $doc -header $header

        $doc.Activate()
        $sel = $word.Selection
        [void]$sel.EndKey(6)   # wdStory = 6

        & $cos $sel

        $doc.Save()
        $doc.Close($false)
    } catch {
        try { $doc.Close($false) } catch { }
        throw
    }
    # Al desti final (xarxa o local). Si no s'hi pot moure, es queda el temporal
    # i es retorna la seva ruta: val mes un informe a %TEMP% que cap informe.
    try { Move-Item -LiteralPath $tempPath -Destination $outPath -Force } catch { return $tempPath }
    return $outPath
}

# ----------------------------------------------------------------------------
# ESCRIURE UNA LINIA DE CATALEG
# ----------------------------------------------------------------------------
# Una linia del cataleg pot portar text i enllacos barrejats ("[[URL]] ..."). El
# motor els separa: el text va com a cos i CADA enllac com a hipervincle en
# paragraf propi.
#
# N'hi havia TRES copies -$emitLine (Document.ps1), _LlicEmetLinia (Llicencia.ps1)
# i _VLine (VistaWord.ps1)-, i les dues primeres nomes es diferenciaven en si
# deduplicaven els enllacos o no.
#
#   -IsChild : sagnia de sub-nivell (el cos i l'enllac d'un fill).
#   $vistos  : conjunt d'enllacos ja emesos EN AQUEST PUNT, per no repetir-los.
#              A Llicencia cal: el text de REQ1 i el comentari "No es disposa..."
#              solen portar el mateix enllac i sortia dues vegades seguides. Amb
#              $null (REQ1) no es dedupa res, que es el comportament de sempre.
#   $emesos  : on s'apunten els enllacos que s'han arribat a escriure. Serveix a
#              _LlicEscriuPunt per saber quins ha de deixar per despres del
#              comentari (l'enllac va DESPRES de la frase que l'anuncia).
#
# La linia ha d'arribar JA RESOLTA (els [CAMP:]/[OPCIO:] es resolen per BLOC,
# no linia a linia: vegeu Apply-FieldsToLines).
#
# ELS ENLLACOS ES DETECTEN AMB _SplitTextAndUrls, MAI A MA. Llicencia va tenir
# un _EsUrl fet amb -like '[[URL]]*', i en un patro de -like '[[URL]' es una
# CLASSE DE CARACTERS: no coincidia mai, el marcador [[URL]] sortia TAL QUAL a
# l'informe i l'enllac no era hipervincle. (Vegeu CLAUDE.md: aquesta trampa ja
# ha sortit tres vegades en aquest projecte.)
function Write-Linia($sel, [string]$linia, [switch]$IsChild, $vistos = $null, $emesos = $null) {
    if ([string]::IsNullOrWhiteSpace($linia)) { return }
    $parts = _SplitTextAndUrls $linia
    if (-not [string]::IsNullOrWhiteSpace($parts.Text)) {
        if ($IsChild) { Format-Body $sel $parts.Text -IsChild } else { Format-Body $sel $parts.Text }
    }
    foreach ($u in @($parts.Urls)) {
        $clau = ([string]$u).Trim()
        if ($null -ne $vistos -and $vistos.Contains($clau)) { continue }
        if ($null -ne $vistos) { [void]$vistos.Add($clau) }
        if ($null -ne $emesos) { [void]$emesos.Add($clau) }
        if ($IsChild) { Format-Url $sel $u -IsChild } else { Format-Url $sel $u }
    }
}

# ----------------------------------------------------------------------------
# EL MOTOR: UN DOCUMENT ES UNA LLISTA DE BLOCS
# ----------------------------------------------------------------------------
# Fins ara, "escriure un punt" (numero + cos + fills + enllacos) estava escrit
# TRES vegades: _WriteCatalegBody (Document.ps1), _LlicEscriuPunt
# (Llicencia.ps1) i _VistaCataleg (VistaWord.ps1). Les dues primeres son el
# mateix algorisme; la tercera tambe, nomes que amb el nivell d'esquema a sobre.
# I cada copia decidia pel seu compte l'aire i quin sub-punt era el primer.
#
# Ara la composicio es parteix en dos:
#   1. un Build-<Familia>Blocs PUR, que diu QUE s'escriu i EN QUIN ORDRE;
#   2. Write-Informe, que ho escriu i es l'UNIC que decideix l'aire, el
#      -First del primer sub-punt i el nivell d'esquema.
#
# El pas 1 es prova a Linux sense Word ni dobles: es una llista de hashtables.
#
# VOCABULARI (la clau 'T' de cada bloc):
#   seccio        Text                    titol de seccio (MAJUSCULES)
#   subseccio     Text                    subseccio (subratllada)
#   etiqueta      Text                    rotul dins del cos
#   item          Num, Text               punt numerat
#   cos           Text, Fill, Negreta, Separat
#   pic           Text, Fill              vinyeta (el -First el posa el motor)
#   nota          Text                    sub-paragraf sagnat sense pic
#   enllac        Url, Fill               hipervincle
#   pla           Text, Negreta, Cos      text pla (ANNEX 1)
#   continua      Text                    segueix el paragraf anterior
#   llista        Text                    paragraf de llista de Word (sol anar buit)
#   conclusio     Text
#   conclusiocap  Text                    el titol CONCLUSIONS
#   aire          Clau                    espai entre blocs SI la bandera ho diu
#   espai                                 linia en blanc SEMPRE (cos fix, ANNEX 1)
#   saltpagina                            pagina nova
#   unitat        Blocs                   UN PUNT SENCER (vegeu mes avall)
#
# 'unitat' es el bloc que fa que aixo funcioni:
#   - obre un punt nou, o sigui que el primer 'pic' de dins es el PRIMER
#     sub-punt i se separa mes de l'item (-First: 12 pt en lloc de 6);
#   - quan es tanca, si ha escrit alguna cosa, hi posa l'aire d'item. Aixi
#     l'espai va despres de l'item COMPLET (amb els seus fills i enllacos), que
#     es el que abans feia el $itemWritten a ma a cada familia.
#
# -AmbNivells: a mes, posa el nivell d'esquema (OutlineLevel) a cada paragraf,
# que es el que fa navegable una VISTA de cataleg al panell del Word. Els
# informes NO el porten, i per aixo es una opcio i no el comportament normal.
$Script:NivellPerTipus = @{
    'seccio'    = 1
    'subseccio' = 2
    'item'      = 3
}

function _MiNivell([string]$tipus) {
    $n = $Script:NivellPerTipus[$tipus]
    if ($null -eq $n) { return $Script:WdOutlineBody }
    return [int]$n
}

function Write-Informe($sel, $blocs, [switch]$AmbNivells) {
    $estat = @{ PrimerFill = $false; Escrits = 0 }
    _WriteBlocs $sel $blocs $estat ([bool]$AmbNivells)
    return $estat.Escrits
}

# El bucle de debo. $estat es un hashtable a posta: els blocs 'unitat' criden
# aquesta funcio recursivament i han de compartir el comptador i el "el proper
# sub-punt es el primer" (un hashtable va per referencia; una variable, no).
function _WriteBlocs($sel, $blocs, $estat, [bool]$ambNivells) {
    foreach ($b in @($blocs)) {
        if ($null -eq $b) { continue }
        $t = [string]$b.T
        $fill = [bool]$b.Fill

        # Els contenidors i l'aire no son paragrafs: van a part.
        if ($t -eq 'unitat') {
            $abans = $estat.Escrits
            $estat.PrimerFill = $true          # el proper sub-punt obre llista
            _WriteBlocs $sel $b.Blocs $estat $ambNivells
            # L'aire va despres de la unitat SENCERA, i nomes si ha escrit res.
            if ($estat.Escrits -gt $abans) { Format-Aire $sel 'item' }
            continue
        }
        # L'aire i els espais son paragrafs BUITS, pero en una vista tambe han de
        # tornar el nivell d'esquema a cos: el Word l'HERETA, i un espaiador
        # despres d'un titol es quedaria a nivell 1 i sortiria com una entrada
        # buida al panell de navegacio.
        if ($t -eq 'aire') {
            Format-Aire $sel ([string]$b.Clau)
            if ($ambNivells) { Format-Nivell $sel $Script:WdOutlineBody }
            continue
        }
        if ($t -eq 'espai') {
            Format-Spacer $sel
            if ($ambNivells) { Format-Nivell $sel $Script:WdOutlineBody }
            continue
        }
        if ($t -eq 'saltpagina') { Format-SaltPagina $sel; continue }

        switch ($t) {
            'seccio'       { Format-Section $sel ([string]$b.Text) }
            'subseccio'    { Format-Subsection $sel ([string]$b.Text) }
            'etiqueta'     { Format-Label $sel ([string]$b.Text) }
            'item'         { Format-Item $sel ([string]$b.Num) ([string]$b.Text) }
            'nota'         { Format-Note $sel ([string]$b.Text) }
            'conclusio'    { Format-Conclusion $sel ([string]$b.Text) }
            'conclusiocap' { Format-ConclusionHeader $sel ([string]$b.Text) }
            'llista'       { Format-ListItem $sel ([string]$b.Text) }
            'continua'     { Format-Append $sel ([string]$b.Text) }
            'cos' {
                if ($fill) { Format-Body $sel ([string]$b.Text) -IsChild -Bold:([bool]$b.Negreta) -Separat:([bool]$b.Separat) }
                else       { Format-Body $sel ([string]$b.Text) -Bold:([bool]$b.Negreta) -Separat:([bool]$b.Separat) }
            }
            'enllac' {
                if ($fill) { Format-Url $sel ([string]$b.Url) -IsChild } else { Format-Url $sel ([string]$b.Url) }
            }
            'pla' {
                Format-Plain $sel ([string]$b.Text) -Bold:([bool]$b.Negreta) -Size ([int]$b.Cos)
            }
            'pic' {
                # EL -First EL DECIDEIX EL MOTOR, no qui munta els blocs: es
                # l'unic que sap si aquest sub-punt obre la llista d'una unitat.
                $primer = [bool]$estat.PrimerFill
                if ($fill) { Format-Bullet $sel ([string]$b.Text) -IsChild -First:$primer }
                else       { Format-Bullet $sel ([string]$b.Text) -First:$primer }
                $estat.PrimerFill = $false
            }
            default { throw ("Write-Informe: tipus de bloc desconegut '" + $t + "'") }
        }

        $estat.Escrits++
        if ($ambNivells) { Format-Nivell $sel (_MiNivell $t) }
    }
}

# ----------------------------------------------------------------------------
# ELS BLOCS D'UN CATALEG (REQ1, TERMINI... i la seva VISTA)
# ----------------------------------------------------------------------------
# Funcio PURA: no toca el Word. Es prova a Linux comptant blocs.
#
# La fan servir _WriteCatalegBody (l'informe) i _VistaCataleg (la vista en
# Word), que abans eren el MATEIX algorisme escrit dues vegades. Amb aixo la
# vista no pot dir una cosa i el document una altra: es el mateix codi.

# Un text de cataleg -> els blocs 'cos' i 'enllac' que li toquen.
function _BlocsDeLinia([string]$linia, [bool]$fill) {
    $out = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($linia)) { return $out.ToArray() }
    $parts = _SplitTextAndUrls $linia
    if (-not [string]::IsNullOrWhiteSpace($parts.Text)) {
        [void]$out.Add(@{ T = 'cos'; Text = $parts.Text; Fill = $fill })
    }
    foreach ($u in @($parts.Urls)) {
        [void]$out.Add(@{ T = 'enllac'; Url = $u; Fill = $fill })
    }
    return $out.ToArray()
}

# Un item del cataleg -> una 'unitat' amb el numero, el cos, els fills i els
# enllacos. $num es un comptador per referencia ([ref]) perque la numeracio va
# seguida de cap a peus del document, no per seccio.
# Les linies d'un node, amb els [CAMP:]/[OPCIO:] resolts o TAL QUAL.
#
# A l'informe es resolen (per BLOC, no linia a linia: un marcador pot ocupar dos
# paragrafs del cataleg). A la VISTA del cataleg NO: alli s'han de veure els
# marcadors, perque es una vista del cataleg i no l'informe d'una activitat
# concreta. Resoldre'ls amb un diccionari buit els deixaria en blanc i la vista
# perdria justament el que hi vas a mirar.
function _LiniesDeNode($node, $fields, [bool]$senseCamps) {
    if ($senseCamps) { return @(@($node.BodyLines) | ForEach-Object { [string]$_ }) }
    return @(Apply-FieldsToLines $node.BodyLines $fields)
}

function _BlocsDItem($el, $fields, [ref]$num, [bool]$senseCamps = $false) {
    $dins = New-Object System.Collections.ArrayList
    $linies = @(_LiniesDeNode $el $fields $senseCamps)
    $fills = @($el.Children)
    $escrit = $false

    if (($el.Selected -or $fills.Count -gt 0) -and $linies.Count -gt 0) {
        $num.Value++
        # Un URL enganxat al text principal de l'item se separa del numero.
        $p0 = _SplitTextAndUrls ([string]$linies[0])
        [void]$dins.Add(@{ T = 'item'; Num = ("$($num.Value)."); Text = $p0.Text })
        foreach ($u in @($p0.Urls)) { [void]$dins.Add(@{ T = 'enllac'; Url = $u }) }
        for ($i = 1; $i -lt $linies.Count; $i++) {
            foreach ($x in @(_BlocsDeLinia ([string]$linies[$i]) $false)) { [void]$dins.Add($x) }
        }
        $escrit = $true
    }

    foreach ($ch in $fills) {
        $cl = @(_LiniesDeNode $ch $fields $senseCamps)
        if ($cl.Count -eq 0) { continue }
        # Un item sense linies propies pero AMB fills tambe consumeix numero.
        if (-not $escrit) { $num.Value++; $escrit = $true }
        # Els fills NO es numeren: van amb pic. El -First el posa el motor.
        $pc = _SplitTextAndUrls ([string]$cl[0])
        if (-not [string]::IsNullOrWhiteSpace($pc.Text)) {
            [void]$dins.Add(@{ T = 'pic'; Text = $pc.Text; Fill = $true })
        }
        foreach ($u in @($pc.Urls)) { [void]$dins.Add(@{ T = 'enllac'; Url = $u; Fill = $true }) }
        for ($i = 1; $i -lt $cl.Count; $i++) {
            foreach ($x in @(_BlocsDeLinia ([string]$cl[$i]) $true)) { [void]$dins.Add($x) }
        }
    }

    if (-not $escrit) { return @() }
    return @(@{ T = 'unitat'; Blocs = $dins.ToArray() })
}

function Build-CatalegBlocs($seccions, $fields, [string]$introText, [bool]$esCosFix = $false, $liniesCosFix = @(), [switch]$SenseCamps) {
    $b = New-Object System.Collections.ArrayList
    $sc = [bool]$SenseCamps

    # Informe de COS FIX (TERMINI): no hi ha seccions ni items a numerar; el cos
    # son els paragrafs del cataleg amb els camps resolts.
    if ($esCosFix) {
        $linies = if ($sc) { @(@($liniesCosFix) | ForEach-Object { [string]$_ }) }
                  else      { @(Apply-FieldsToLines $liniesCosFix $fields) }
        for ($i = 0; $i -lt $linies.Count; $i++) {
            $l = [string]$linies[$i]
            if ([string]::IsNullOrWhiteSpace($l)) { continue }
            foreach ($x in @(_BlocsDeLinia $l $false)) { [void]$b.Add($x) }
            # SEMPRE, no per bandera: aqui la linia en blanc es part del text
            # del cataleg (son paragrafs solts, no blocs d'un informe).
            if ($i -lt ($linies.Count - 1)) { [void]$b.Add(@{ T = 'espai' }) }
        }
        return $b.ToArray()
    }

    if (-not [string]::IsNullOrWhiteSpace($introText)) {
        [void]$b.Add(@{ T = 'cos'; Text = $introText })
        [void]$b.Add(@{ T = 'aire'; Clau = 'introparagraf' })
    }

    $num = 0
    $darreraSeccio = $null
    foreach ($sec in @($seccions)) {
        # "Seccio - Subseccio" (el que munta Build-SelectionFromKeys) o el titol
        # sol (el que arriba del cataleg sencer, a la vista).
        $parts = ([string]$sec.Title) -split ' - ', 2
        if ($parts.Count -eq 2) {
            $nomSec = $parts[0].Trim()
            if ($nomSec -ne $darreraSeccio) {
                [void]$b.Add(@{ T = 'seccio'; Text = $nomSec })
                [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
                $darreraSeccio = $nomSec
            }
            [void]$b.Add(@{ T = 'subseccio'; Text = $parts[1].Trim() })
            [void]$b.Add(@{ T = 'aire'; Clau = 'subseccio' })
        } else {
            [void]$b.Add(@{ T = 'seccio'; Text = [string]$sec.Title })
            [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
            $darreraSeccio = [string]$sec.Title
        }

        # Les subseccions i les intros s'emeten TARD: nomes quan ve un item de
        # debo que les segueix. Si una seccio te 3 subseccions i nomes s'ha
        # triat un item de la tercera, no han de sortir les dues primeres
        # buides.
        #
        # UN TEXT FIX ES DE LA SECCIO O DE LA SUBSECCIO, SEGONS ON ESTIGUI.
        # Abans, QUALSEVOL subseccio invalidava l'intro pendent. Amb aixo, un
        # text posat a la SECCIO -abans de la primera subseccio- no sortia MAI:
        # els seus items pengen de les subseccions, i el primer marcador de
        # subseccio ja se l'havia endut. Va passar de debo en moure el text de
        # l'article 4 de l'Ordenanca de dins de "Certificats d'inscripcio en el
        # RITSIC" a la seccio "Instal-lacions", que es on toca perque parla
        # tambe de les inspeccions: el text va DESAPAREIXER de tots els
        # informes, i nomes es va veure al fitxer d'or.
        #   - intro d'ABANS de la primera subseccio -> es de la SECCIO: sobreviu
        #     als canvis de subseccio i surt amb el PRIMER item que s'emeti.
        #   - intro de DINS d'una subseccio -> es d'aquella subseccio i mor amb
        #     ella (que es el que evitava que sortis damunt d'una altra).
        $subPendent = $null
        $introPendent = $null
        $introSeccio = $null
        $dinsSub = $false
        foreach ($el in @($sec.Items)) {
            if ([string]$el.Kind -eq 'subsection') {
                $subPendent = $el; $dinsSub = $true; $introPendent = $null; continue
            }
            if ([string]$el.Kind -eq 'intro') {
                if ($dinsSub) { $introPendent = $el } else { $introSeccio = $el }
                continue
            }

            $blocsItem = @(_BlocsDItem $el $fields ([ref]$num) $sc)
            if ($blocsItem.Count -eq 0) { continue }

            # L'intro de la SECCIO va abans del titol de la subseccio: introdueix
            # tot el que ve despres, no nomes el primer grup.
            if ($null -ne $introSeccio) {
                foreach ($ln in @(_LiniesDeNode $introSeccio $fields $sc)) {
                    foreach ($x in @(_BlocsDeLinia ([string]$ln) $false)) { [void]$b.Add($x) }
                }
                [void]$b.Add(@{ T = 'aire'; Clau = 'intro' })
                $introSeccio = $null
            }
            if ($null -ne $subPendent) {
                [void]$b.Add(@{ T = 'subseccio'; Text = [string]$subPendent.Short })
                [void]$b.Add(@{ T = 'aire'; Clau = 'subseccio' })
                $subPendent = $null
            }
            if ($null -ne $introPendent) {
                foreach ($ln in @(_LiniesDeNode $introPendent $fields $sc)) {
                    foreach ($x in @(_BlocsDeLinia ([string]$ln) $false)) { [void]$b.Add($x) }
                }
                [void]$b.Add(@{ T = 'aire'; Clau = 'intro' })
                $introPendent = $null
            }
            foreach ($x in $blocsItem) { [void]$b.Add($x) }
        }
    }
    return $b.ToArray()
}
