#requires -Version 5.1
<#
.SYNOPSIS
  Llicencia: QUE s'escriu a l'informe. Funcions PURES que tornen una llista de
  blocs (el vocabulari de Write-Informe, MotorInforme.ps1) i Build-LlicenciaDocument,
  que els escriu. La regla de seccions, subseccions i textos fixos es
  _LlicBlocsPunts i la comparteix la vista de LLIC.json (VistaWord.ps1).
  Vegeu la capcalera de LlicenciaDades.ps1 per al mapa del modul.
#>

# Cos de lletra del full de signatures de l'ANNEX 1 (a la plantilla, sz=18
# mig-punts = 9 pt). La resta de l'informe va a 11.
$Script:LlicAnnexSignaturaCos = 9

# El titol que obre el full de signatures de l'ANNEX 1. Es mira pel PRINCIPI del
# text -es una frase llarga amb citacions legals- i sense accents, per no
# dependre de com s'hagi escrit al cataleg. Funcio PURA.
function _LlicEsTitolAcceptacio([string]$text) {
    $t = ([string]$text).Trim().ToLower()
    if ($t.Length -lt 12) { return $false }
    $t = $t.Replace([char]0x2019, "'").Replace([char]0x00F3, 'o').Replace([char]0x00E9, 'e')
    return $t.StartsWith("document d'acceptacio")
}

# ----------------------------------------------------------------------------
# COMPOSICIO DEL DOCUMENT: BLOCS PURS + Write-Informe (MotorInforme.ps1)
# ----------------------------------------------------------------------------
# Llicencia munta el document com REQ1: unes funcions PURES diuen QUE s'escriu
# i en quin ordre (una llista de blocs, que es prova a Linux sense Word) i
# Write-Informe ho escriu. Abans ho escrivia tot pel seu compte, i la regla de
# les seccions, subseccions i textos fixos era a TRES llocs -REQ1, aquest
# informe i la vista en Word de LLIC.json- que s'havien de tocar alhora (un
# text fix va arribar a sortir tres vegades). Ara la de Llicencia es UNA
# (_LlicBlocsPunts) i la fan servir l'informe i la vista.

# EL TEXT DEL "Quan:". Funcio PURA.
#
# ENTRE PARENTESIS i sense el punt final, a les TRES fases:
# "(Quan: Abans d'iniciar l'activitat)" -decisio de l'usuari, setembre 2026;
# primer nomes al pre i despres a tots-. El $fase es queda per si algun dia
# ha de tornar a ser diferent en alguna fase.
function _LlicTextQuan([string]$quan, [string]$fase = '') {
    $q = ([string]$quan).Trim()
    return ('(Quan: ' + $q.TrimEnd('.').TrimEnd() + ')')
}

# UN PUNT DE LLICENCIA, en blocs (el que va DINS de la 'unitat'): el cos (de
# REQ1 o propi) com a item numerat, els sub-punts amb pic, el "(Quan: ...)" i
# el comentari "No es disposa..." / "Es disposa...". Funcio PURA.
#
# ON VA L'ENLLAC. El comentari acaba dient "...en el seguent enllac:", o sigui
# que l'enllac ha d'anar JUST DESPRES d'aquella frase. Pero el cos de l'item (de
# REQ1) sol portar EL MATEIX enllac, i sortia abans -amb la frase penjada sense
# res al darrere-. Per aixo es miren PRIMER els enllacos del comentari: els que
# tambe son al cos NO s'emeten amb l'item; surten despres del comentari. Cap
# enllac es repeteix dins d'un punt.
#
# ELS CAMPS ES RESOLEN PER BLOC (Apply-FieldsToLines) i NO linia a linia: un
# [OPCIO:]/[CAMP:] pot ocupar dos paragrafs del cataleg.
function _LlicBlocsDePunt($punt, [string]$marca, $fields, [string]$estat, [bool]$ambQuan, [string]$fase = '') {
    $out = New-Object System.Collections.ArrayList
    $comLinies = if ($estat -eq 'si') { @(Apply-FieldsToLines $punt.SiDisposa $fields) }
                 elseif ($estat -eq 'no') { @(Apply-FieldsToLines $punt.NoDisposa $fields) }
                 else { @() }
    # Els enllacos ja emesos (o reservats per al comentari) en AQUEST punt.
    $vistos = New-Object System.Collections.ArrayList
    foreach ($l in $comLinies) {
        foreach ($u in @((_SplitTextAndUrls ([string]$l)).Urls)) {
            $c = ([string]$u).Trim()
            if (-not $vistos.Contains($c)) { [void]$vistos.Add($c) }
        }
    }
    $emesos = New-Object System.Collections.ArrayList
    # Un enllac que encara no ha sortit: s'apunta i es torna el bloc.
    $enllac = {
        param($u, [bool]$fill)
        $c = ([string]$u).Trim()
        if ($vistos.Contains($c)) { return }
        [void]$vistos.Add($c); [void]$emesos.Add($c)
        [void]$out.Add(@{ T = 'enllac'; Url = $u; Fill = $fill })
    }

    $linies = @(Apply-FieldsToLines $punt.Cos $fields)
    $primera = if ($linies.Count -gt 0) { [string]$linies[0] } else { [string](Apply-Fields -text $punt.Titol -fields $fields) }
    # El numero i el text van junts a l'item; l'URL de la PRIMERA linia, a part.
    $p0 = _SplitTextAndUrls ([string]$primera)
    [void]$out.Add(@{ T = 'item'; Num = [string]$marca; Text = $p0.Text })
    foreach ($u in @($p0.Urls)) { & $enllac $u $false }
    for ($i = 1; $i -lt $linies.Count; $i++) {
        $pp = _SplitTextAndUrls ([string]$linies[$i])
        if (-not [string]::IsNullOrWhiteSpace($pp.Text)) { [void]$out.Add(@{ T = 'cos'; Text = $pp.Text }) }
        foreach ($u in @($pp.Urls)) { & $enllac $u $false }
    }
    # Sub-punts. L'ENLLAC D'UN SUB-PUNT SENSE TEXT NO VA SAGNAT: no penja de cap
    # pic visible, i sagnat quedava despenjat un centimetre a la dreta. El -First
    # del primer pic el posa el motor (la 'unitat').
    foreach ($sub in @($punt.Subs)) {
        $ambPic = $false
        foreach ($l in @(Apply-FieldsToLines $sub $fields)) {
            $pc = _SplitTextAndUrls ([string]$l)
            if (-not [string]::IsNullOrWhiteSpace($pc.Text)) {
                [void]$out.Add(@{ T = 'pic'; Text = $pc.Text; Fill = $true })
                $ambPic = $true
            }
            foreach ($u in @($pc.Urls)) { & $enllac $u $ambPic }
        }
    }
    # EL "Quan:" VA ABANS DEL COMENTARI: a l'informe fet a ma, primer QUAN s'ha
    # de tenir i despres SI ES TE.
    if ($ambQuan) {
        foreach ($l in @(Apply-FieldsToLines $punt.Quan $fields)) {
            if ([string]::IsNullOrWhiteSpace($l)) { continue }
            [void]$out.Add(@{ T = 'cos'; Text = (_LlicTextQuan ([string]$l) $fase) })
        }
    }
    # El comentari: 'no' = falta (negreta a la primera linia); 'si' = ja hi es.
    # La PRIMERA linia va SEPARADA (els 12 pt que separen un item del seu primer
    # sub-punt), com a l'informe fet a ma. Els seus enllacos surten aqui, que es
    # el lloc que la frase anuncia.
    $primerCom = $true
    foreach ($l in $comLinies) {
        $pp = _SplitTextAndUrls ([string]$l)
        if (-not [string]::IsNullOrWhiteSpace($pp.Text)) {
            [void]$out.Add(@{ T = 'cos'; Text = $pp.Text; Negreta = ($primerCom -and $estat -eq 'no'); Separat = $primerCom })
            $primerCom = $false
        }
        foreach ($u in @($pp.Urls)) {
            $c = ([string]$u).Trim()
            if ($emesos.Contains($c)) { continue }
            [void]$emesos.Add($c)
            [void]$out.Add(@{ T = 'enllac'; Url = $u })
        }
    }
    return $out.ToArray()
}

# EL BLOC PROJECTE I LA DOCUMENTACIO DEL PROJECTE SON COMPLEMENTARIS, i les dues
# regles van juntes perque son la mateixa idea vista de les dues bandes:
#
#   REQUERIMENT   -> hi ha requeriments de projecte (bloc PROJECTE, amb lletres)
#                    i per tant el projecte i els planols encara s'han de
#                    modificar: la documentacio NO es definitiva i no hi surt.
#   FAVORABLES    -> ja no queden requeriments de projecte (el bloc no hi surt)
#                    i la documentacio ja es la bona: hi va, i dalt de tot.
#
# Funcions PURES, i la regla escrita en un sol lloc.
function _LlicPortaDocProjecte([string]$fase) {
    return ([string]$fase -eq 'favorable-pre' -or [string]$fase -eq 'favorable-post')
}

function _LlicPortaProjecte([string]$fase) {
    return ([string]$fase -eq 'requeriment')
}

# LA MARCA D'UN PUNT: "1." o "A.". PURA. La lletra la fa _LlicLletra
# (LlicenciaDades.ps1): tambe dona nom als adjunts, i si fos aqui les dades
# dependrien de la composicio.
function _LlicMarca([int]$i, [string]$estil) {
    if ([string]$estil -eq 'lletra') { return ((_LlicLletra $i) + '.') }
    return ([string]$i + '.')
}

# UNA LLISTA DE PUNTS AMB LA SEVA ESTRUCTURA DE REQ1, en blocs: la SECCIO en
# majuscules, la SUBSECCIO i el TEXT FIX que l'encapcala, i despres els punts
# (cada un en una 'unitat', que hi posa l'aire d'item). Funcio PURA.
#
# LA REGLA ES AQUI I NOMES AQUI: la fan servir l'informe (Build-LlicenciaBlocs)
# i la vista de LLIC.json (_VistaLlicencia). El que canvia es COM s'escriu cada
# punt, i per aixo arriba com a scriptblock: & $blocsDePunt $punt $marca.
#
# La capcalera nomes s'escriu quan CANVIA, i com que penja de cada punt, una
# seccio sense punts no pot sortir. L'INTRO SURT QUAN CANVIA (no quan canvia la
# subseccio): aixi surt amb el PRIMER punt del grup encara que no sigui el
# primer del cataleg, i un grup amb dos intros els treu tots dos. L'intro de la
# SECCIO va abans del titol de la subseccio i no es repeteix a cada subseccio.
#
# $n va per REFERENCIA: la numeracio de l'informe es SEGUIDA de cap a peus.
# $senseCamps: la VISTA ensenya els [CAMP:] tal qual.
function _LlicBlocsPunts($punts, [ref]$n, $fields, [string]$estil, [scriptblock]$blocsDePunt, [bool]$senseCamps = $false) {
    $b = New-Object System.Collections.ArrayList
    $secAra = $null; $subAra = $null
    $introAra = $null      # intro d'una SUBSECCIO: es reinicia a cada grup
    $introSecAra = $null   # intro de la SECCIO: nomes es reinicia amb la seccio
    foreach ($p in @($punts)) {
        $sec = [string]$p.Seccio
        $sb  = [string]$p.Subseccio
        # NOMES si te text de debo: una linia buida faria sortir l'aire sol.
        $crues = if ($senseCamps) { @(@($p.Intro) | ForEach-Object { [string]$_ }) } else { @(Apply-FieldsToLines @($p.Intro) $fields) }
        $intro = @($crues | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        $clauIntro = ($intro -join "`n")
        $deSeccio = [bool]$p.IntroDeSeccio

        if ($sec -ne $secAra) {
            if (-not [string]::IsNullOrWhiteSpace($sec)) {
                [void]$b.Add(@{ T = 'seccio'; Text = $sec })
                [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
            }
            $secAra = $sec; $subAra = $null; $introAra = $null; $introSecAra = $null
        }
        if ($deSeccio -and $intro.Count -gt 0 -and $clauIntro -ne $introSecAra) {
            foreach ($l in $intro) { foreach ($x in @(_BlocsDeLinia ([string]$l) $false)) { [void]$b.Add($x) } }
            [void]$b.Add(@{ T = 'aire'; Clau = 'intro' })
            $introSecAra = $clauIntro
        }
        if ($sb -ne $subAra) {
            if (-not [string]::IsNullOrWhiteSpace($sb)) {
                [void]$b.Add(@{ T = 'subseccio'; Text = $sb })
                [void]$b.Add(@{ T = 'aire'; Clau = 'subseccio' })
            }
            $subAra = $sb; $introAra = $null
        }
        if (-not $deSeccio -and $intro.Count -gt 0 -and $clauIntro -ne $introAra) {
            foreach ($l in $intro) { foreach ($x in @(_BlocsDeLinia ([string]$l) $false)) { [void]$b.Add($x) } }
            [void]$b.Add(@{ T = 'aire'; Clau = 'intro' })
            $introAra = $clauIntro
        }
        $n.Value++
        [void]$b.Add(@{ T = 'unitat'; Blocs = @(& $blocsDePunt $p (_LlicMarca $n.Value $estil)) })
    }
    return $b.ToArray()
}

# L'INFORME SENCER en blocs. Funcio PURA.
#
# $model porta tot el que ha triat l'usuari a l'assistent:
#   Fase, EsProvisional, Header, Fields, Abans, Projecte, Despres, Doc,
#   CondicionsActors (noms dels actors marcats), Cataleg.
function Build-LlicenciaBlocs($model) {
    $b = New-Object System.Collections.ArrayList
    $fields = $model.Fields
    $fase = [string]$model.Fase

    # ---- DOCUMENTACIO DEL PROJECTE ----
    # NOMES ALS FAVORABLES, i alli al principi de tot. Al REQUERIMENT no hi va:
    # s'hi demanen modificacions al projecte i als planols, o sigui que aquella
    # documentacio ENCARA NO ES DEFINITIVA. (Decisio de l'usuari, agost 2026.)
    if (_LlicPortaDocProjecte $fase) {
        $doc1 = [string]$model.Doc.Text
        if (-not [string]::IsNullOrWhiteSpace($doc1)) {
            [void]$b.Add(@{ T = 'titolbloc'; Text = ('DOCUMENTACI' + [char]0x00D3 + ' PROJECTE') })
            [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
            $u = New-Object System.Collections.ArrayList
            [void]$u.Add(@{ T = 'cos'; Text = $doc1 })
            foreach ($d in @($model.Doc.Items)) { [void]$u.Add(@{ T = 'pic'; Text = [string]$d; Fill = $true }) }
            [void]$b.Add(@{ T = 'unitat'; Blocs = $u.ToArray() })
        }
    }

    # Com s'escriu un punt d'aquest informe (per a _LlicBlocsPunts). SENSE
    # .GetNewClosure(): s'invoca dins d'aquesta mateixa crida i veu $fields i
    # $fase en temps d'execucio (el mateix patro que Write-InformeDocx).
    $puntInforme = { param($p, $marca) _LlicBlocsDePunt $p $marca $fields ([string]$p.Estat) $false $fase }

    # ---- PROJECTE ----
    # NOMES AL REQUERIMENT (als favorables ja estan resolts). Va el PRIMER i amb
    # LLETRES: quan quedin resolts, el bloc desapareix i la resta ha de
    # conservar la MATEIXA numeracio (el titular ho compara amb el que ja tenia).
    $proj = if (_LlicPortaProjecte $fase) { @($model.Projecte) } else { @() }
    if ($proj.Count -gt 0) {
        [void]$b.Add(@{ T = 'titolbloc'; Text = (_LlicTitolProjecte) })
        [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
        $lletra = 0
        foreach ($x in @(_LlicBlocsPunts $proj ([ref]$lletra) $fields 'lletra' $puntInforme)) { [void]$b.Add($x) }
    }

    # ---- ABANS ----
    [void]$b.Add(@{ T = 'titolbloc'; Text = (_LlicTitolAbans) })
    [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
    $n = 0
    foreach ($x in @(_LlicBlocsPunts @($model.Abans) ([ref]$n) $fields 'numero' $puntInforme)) { [void]$b.Add($x) }

    # ---- DESPRES ----
    # LA NUMERACIO CONTINUA la del bloc ABANS ($n no es reinicia). El bloc
    # PROJECTE no hi compta: te el seu comptador de lletres.
    $desp = @($model.Despres)
    if ($desp.Count -gt 0) {
        [void]$b.Add(@{ T = 'titolbloc'; Text = (_LlicTitolDespres) })
        [void]$b.Add(@{ T = 'aire'; Clau = 'seccio' })
        $puntDespres = { param($p, $marca) _LlicBlocsDePunt $p $marca $fields ([string]$p.Estat) $true $fase }
        foreach ($x in @(_LlicBlocsPunts $desp ([ref]$n) $fields 'numero' $puntDespres)) { [void]$b.Add($x) }
    }

    # ---- CONCLUSIO ----
    # Mateix bloc que REQ1: capcalera CONCLUSIONS centrada i en negreta, i la
    # conclusio en negreta (el **...** ve del cataleg). Amb condicions, a sota
    # els ACTORS amb lletres minuscules (a., b., c.), com el Word de l'usuari.
    $actorsCond = @(@($model.CondicionsActors) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $ambCond = ($actorsCond.Count -gt 0 -and (_LlicAdmetCondicions $fase))
    [void]$b.Add(@{ T = 'aire'; Clau = 'conclusions' })
    [void]$b.Add(@{ T = 'conclusiocap'; Text = 'CONCLUSIONS' })
    [void]$b.Add(@{ T = 'conclusio'; Text = (_LlicConclusioText $fase $ambCond) })
    if ($ambCond) {
        for ($i = 0; $i -lt $actorsCond.Count; $i++) {
            $item = @{ T = 'item'; Num = ((_LlicLletra ($i + 1)).ToLower() + '.'); Text = [string]$actorsCond[$i] }
            [void]$b.Add(@{ T = 'unitat'; Blocs = @($item) })
        }
    }
    # El tancament: del cataleg, com tots els altres informes.
    foreach ($x in @(_BlocsTancament $fields)) { [void]$b.Add($x) }

    # ---- ANNEX 1: nomes al REQUERIMENT d'una llicencia PROVISIONAL ----
    if ($fase -eq 'requeriment' -and (_LlicCalAnnex1 $model.Abans ([bool]$model.EsProvisional))) {
        foreach ($x in @(_LlicBlocsAnnex1 $model.Cataleg)) { [void]$b.Add($x) }
    }
    return $b.ToArray()
}

# Composa l'informe sencer i el desa. Retorna la ruta.
function Build-LlicenciaDocument($word, $model) {
    $header = $model.Header
    $baseName = _LlicNomFitxer (Get-Date) ([string]$model.Fase) ([string]$header['ID_GIA'])
    $blocs = Build-LlicenciaBlocs $model
    # La capcalera de LLICENCIA (porta la linia "Classificacio:"). Si el bloc no
    # hi es (0 CAPCALERA.docx encara sense actualitzar), Select-CapcaleraBlock es
    # queda amb el generic i l'informe surt igualment, sense la classificacio.
    return Write-InformeDocx $word $baseName 'LLIC' $header {
        param($sel)
        [void](Write-Informe $sel $blocs)
    }
}

# L'ANNEX 1, que nomes va al REQUERIMENT d'una llicencia provisional. El text es
# FIX i viu al cataleg (seccio que comenca per "ANNEX 1"), no encastat aqui:
# aixi l'usuari el pot editar des de l'editor de catalegs com tota la resta.
function _LlicSeccioAnnex1($llic) {
    foreach ($s in @($llic.nodes)) {
        if ([string]$s.titol -like 'ANNEX 1*') { return $s }
    }
    return $null
}

# L'ANNEX 1 en blocs. Funcio PURA.
#
# Va en TEXT PLA ('pla'): sense sagnies, sense pics i sense numeracio (a la
# plantilla de l'usuari es text corrent, encara que alli el Word hi tingues una
# llista). NOMES van en negreta els dos TITOLS:
#   - "ANNEX 1. Documentacio per demanar..."
#   - "Document d'acceptacio del cessament dels usos..."
# I des del "Document d'acceptacio..." fins al final: PAGINA NOVA i cos 9 (a la
# plantilla, sz=18 mig-punts).
function _LlicBlocsAnnex1($llic) {
    $b = New-Object System.Collections.ArrayList
    $sec = _LlicSeccioAnnex1 $llic
    if ($null -eq $sec) { return $b.ToArray() }
    # Salt de pagina: l'annex es un document a part dins de l'informe.
    [void]$b.Add(@{ T = 'saltpagina' })
    [void]$b.Add(@{ T = 'pla'; Text = [string]$sec.titol; Negreta = $true })

    $cos9 = $false          # ja som al full de signatures?
    $num = 0                # el comptador dels punts numerats
    $primerItem = $true     # el primer punt no porta linia en blanc al davant
    $obreBloc = $false      # al full de signatures, comenca una declaracio nova
    foreach ($nd in @($sec.fills)) {
        # LA MARCA VA COM A TEXT, no com a llista del Word: l'usuari la vol
        # PLANA. El comptador NOMES avanca amb els 'item'.
        $marca = ''
        $tip = [string]$nd.tipus
        if ($tip -eq 'item') { $num++; $marca = [string]$num + '. ' }
        elseif ($tip -eq 'subitem') { $marca = '- ' }
        # UNA LINIA EN BLANC entre punts numerats (a l'informe fet a ma n'hi ha).
        if ($tip -eq 'item' -and -not $primerItem -and -not $cos9) { [void]$b.Add(@{ T = 'espai' }) }
        if ($tip -eq 'item') { $primerItem = $false }
        $primeraLinia = $true
        foreach ($l in @(_LlicCos $nd)) {
            $pp = _SplitTextAndUrls ([string]$l)
            $t = [string]$pp.Text
            if (-not [string]::IsNullOrWhiteSpace($t)) {
                $esTitolAcceptacio = (_LlicEsTitolAcceptacio $t)
                if ($esTitolAcceptacio -and -not $cos9) {
                    # A partir d'aqui, full a part i lletra mes petita.
                    [void]$b.Add(@{ T = 'saltpagina' })
                    $cos9 = $true
                    $obreBloc = $false
                }
                if ($cos9) {
                    # EL FULL DE SIGNATURES: dues linies en blanc davant de cada
                    # declaracio (despres del titol i de cada "Signat").
                    if ($obreBloc) { [void]$b.Add(@{ T = 'espai' }); [void]$b.Add(@{ T = 'espai' }); $obreBloc = $false }
                    [void]$b.Add(@{ T = 'pla'; Text = $t; Negreta = $esTitolAcceptacio; Cos = $Script:LlicAnnexSignaturaCos })
                    if ($esTitolAcceptacio -or $t.Trim() -eq 'Signat') { $obreBloc = $true }
                    $primeraLinia = $false
                }
                elseif ($tip -eq 'text' -and -not $primerItem) {
                    # UN NODE 'text' ES LA CONTINUACIO del punt de sobre: a la
                    # plantilla va DINS del mateix paragraf, separat per un espai.
                    [void]$b.Add(@{ T = 'continua'; Text = (' ' + $t) })
                }
                else {
                    # La marca nomes a la PRIMERA linia del punt.
                    $txt = if ($primeraLinia) { $marca + $t } else { $t }
                    [void]$b.Add(@{ T = 'pla'; Text = $txt })
                    $primeraLinia = $false
                }
            }
            foreach ($u in @($pp.Urls)) { [void]$b.Add(@{ T = 'enllac'; Url = $u }) }
        }
    }
    return $b.ToArray()
}
