#requires -Version 5.1
<#
.SYNOPSIS
  Informe de LLICENCIA d'activitat (Annex II de la Llei 20/2009 i llicencia
  provisional). Tres informes encadenats sobre el mateix expedient.

.DESCRIPTION
  El full de ruta d'una llicencia son TRES informes:

    1. REQUERIMENT           "Cal requerir l'esmena de les deficiencies..."
    2. FAVORABLE PRE         "S'informa favorablement a l'espera de rebre la
                              citada documentacio..." (+ condicions, opcional)
    3. FAVORABLE POST        "S'informa favorablement l'activitat i es dona per
                              tancat l'expedient."

  El primer es opcional, pero es fa gairebe sempre.

  DIFERENCIA AMB UN REQUERIMENT NORMAL: aqui els punts no son deficiencies sino
  DOCUMENTACIO, i surten TANT si es te com si no. Per cada punt s'hi tria:
    - "No es disposa..."  -> NEGRETA (falta)
    - "Es disposa... (Id Firmadoc: ...)" -> sense negreta (ja hi es)
  Al Word que feia servir l'usuari sortien en verd, pero el color era una MARCA
  SEVA per veure que havia de canviar a cada informe; al document generat van
  amb el color de sempre (Format.ps1).

  D'ON SURT EL TEXT: el cos de cada punt es de REQ1, EN VIU. LLIC.json nomes hi
  afegeix el que es propi de Llicencia (els dos comentaris i el "Quan:") i una
  CLAU que apunta a l'item de REQ1 ("Seccio::Titol"). Aixi, canviar un text a
  REQ1 el canvia tambe aqui i no hi ha dues copies per mantenir. Els punts que
  no tenen equivalent a REQ1 porten el text a LLIC (no duen clau).

  ESTRUCTURA DE L'INFORME:
    DOCUMENTACIO NECESSARIA ABANS DE LA RESOLUCIO...
      (punt condicional segons Annex II / llicencia provisional)
      Autoritzacions / Informes preceptius   <- bloc ABANS de LLIC
      Projecte                               <- requeriments normals de REQ1
      Documentacio                           <- tecnic redactor + Id Firmadoc
    DOCUMENTACIO NECESSARIA DESPRES DE LA RESOLUCIO... (amb "Quan:")
    Conclusio de la fase
    ANNEX 1   <- nomes si REQUERIMENT i llicencia provisional

  Les funcions de dades son PURES (es proven en headless, sense Word); nomes
  l'assistent i la composicio del document fan servir WinForms i Word COM.
#>

# ----------------------------------------------------------------------------
# DEFINICIONS (un sol lloc)
# ----------------------------------------------------------------------------
# Les tres fases. 'Clau' es el que es desa i es compara; 'Nom' el que es
# veu; 'Conclusio' el text que tanca l'informe.
function _LlicFases {
    return @(
        [pscustomobject]@{
            Clau = 'requeriment'
            Nom  = 'Requeriment'
            Sub  = 'Es demana la documentaci' + [char]0x00F3 + ' que falta'
            Conclusio = ('Cal requerir l' + [char]0x2019 + 'esmena de les defici' + [char]0x00E8 +
                         'ncies indicades, aportant la documentaci' + [char]0x00F3 + ' corresponent.')
        }
        [pscustomobject]@{
            Clau = 'favorable-pre'
            Nom  = 'Favorable pre-llic' + [char]0x00E8 + 'ncia'
            Sub  = 'Ja hi ha tota la documentaci' + [char]0x00F3 + ' d' + [char]0x2019 + 'abans de la resoluci' + [char]0x00F3
            Conclusio = ('S' + [char]0x2019 + 'informa favorablement a l' + [char]0x2019 + 'espera de rebre la citada ' +
                         'documentaci' + [char]0x00F3 + ' en els terminis de temps especificats per poder donar per ' +
                         'tancat l' + [char]0x2019 + 'expedient')
        }
        [pscustomobject]@{
            Clau = 'favorable-post'
            Nom  = 'Favorable post-llic' + [char]0x00E8 + 'ncia'
            Sub  = 'Ja s' + [char]0x2019 + 'ha comprovat tota la documentaci' + [char]0x00F3
            Conclusio = ('S' + [char]0x2019 + 'informa favorablement l' + [char]0x2019 + 'activitat i es d' +
                         [char]0x00F3 + 'na per tancat l' + [char]0x2019 + 'expedient.')
        }
    )
}

# Titols dels dos grans blocs de l'informe.
function _LlicTitolAbans {
    return ('DOCUMENTACI' + [char]0x00D3 + ' NECESS' + [char]0x00C0 + 'RIA ABANS DE LA RESOLUCI' + [char]0x00D3 +
            ' DE L' + [char]0x2019 + [char]0x00D2 + 'RGAN T' + [char]0x00C8 + 'CNIC AMBIENTAL.')
}
function _LlicTitolDespres {
    return ('DOCUMENTACI' + [char]0x00D3 + ' NECESS' + [char]0x00C0 + 'RIA DESPR' + [char]0x00C9 + 'S DE LA RESOLUCI' +
            [char]0x00D3 + ' DE L' + [char]0x2019 + [char]0x00D2 + 'RGAN T' + [char]0x00C8 + 'CNIC AMBIENTAL EN ELS ' +
            'TERMINIS DE TEMPS ESPECIFICATS.')
}

# ----------------------------------------------------------------------------
# FUNCIONS PURES (testejables en headless)
# ----------------------------------------------------------------------------

# Ruta del cataleg de Llicencia.
function _LlicCatalegPath {
    return [string](Join-Path $EstructuralsDir 'LLIC.json')
}

# Llegeix LLIC.json. Retorna $null si no hi es (el programa ha de dir-ho, no
# fer com si res).
function Read-LlicCataleg([string]$path = '') {
    if ([string]::IsNullOrWhiteSpace($path)) { $path = _LlicCatalegPath }
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try { return (Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}

# Aplana el cos d'un node del JSON a les mateixes linies amb marques que fa
# servir tot el programa (**negreta**, //cursiva//, [[URL]]...). Reaprofita
# _JsonParaToBodyLine de CatalegJson.ps1: el format es el mateix.
function _LlicCos($node) {
    $out = New-Object System.Collections.ArrayList
    foreach ($p in @($node.cos)) { [void]$out.Add((_JsonParaToBodyLine $p)) }
    return $out.ToArray()
}

# El fill d'un item de LLIC amb aquell tipus ('nodisposa', 'sidisposa', 'quan',
# 'subitem'), o $null. Els tipus son els del format estandard, ampliat.
function _LlicFill($item, [string]$tipus) {
    foreach ($f in @($item.fills)) {
        if ([string]$f.tipus -eq $tipus) { return $f }
    }
    return $null
}

# Tots els fills d'un tipus (per als sub-punts, que poden ser-ne uns quants).
function _LlicFills($item, [string]$tipus) {
    $out = New-Object System.Collections.ArrayList
    foreach ($f in @($item.fills)) {
        if ([string]$f.tipus -eq $tipus) { [void]$out.Add($f) }
    }
    return $out.ToArray()
}

# Index dels items de REQ1 per clau ("Seccio::Titol"), per poder-hi anar de
# pressa. $parsed es el que retorna Get-ParsedCataleg / Read-CatalegJson.
function _LlicIndexReq1($parsed) {
    $idx = @{}
    if ($null -eq $parsed) { return $idx }
    foreach ($sec in @($parsed.Sections)) {
        foreach ($el in @($sec.Items)) {
            if ([string]$el.Kind -ne 'item') { continue }
            if ([string]::IsNullOrWhiteSpace([string]$el.Short)) { continue }
            $idx[(_ItemKey $sec.Title $el.Short)] = $el
        }
    }
    return $idx
}

# EL COS DE L'EINA, i es PUR: resol els punts d'un bloc de LLIC ('ABANS',
# 'DESPRES' o 'PROPIS') ajuntant-los amb el text de REQ1.
#
# Retorna @{ Punts; Orfes }:
#   Punts : llista de @{ Clau; Titol; Cos; NoDisposa; SiDisposa; Quan; Subs;
#                        Condicio } en l'ordre del cataleg.
#   Orfes : claus que son a LLIC pero JA NO a REQ1. NO s'amaguen: si el lligam
#           s'ha trencat (perque algu ha reanomenat un requeriment), el
#           programa ho ha de dir en lloc de deixar-se un punt en silenci.
function _LlicPuntsPerBloc($llic, $idxReq1, [string]$bloc) {
    $punts = New-Object System.Collections.ArrayList
    $orfes = New-Object System.Collections.ArrayList
    if ($null -eq $llic) { return @{ Punts = $punts.ToArray(); Orfes = $orfes.ToArray() } }

    $sec = $null
    foreach ($s in @($llic.nodes)) {
        if ([string]$s.titol -eq $bloc) { $sec = $s; break }
    }
    if ($null -eq $sec) { return @{ Punts = $punts.ToArray(); Orfes = $orfes.ToArray() } }

    foreach ($it in @($sec.fills)) {
        $clau = [string]$it.clau
        $cos = @()
        if (-not [string]::IsNullOrWhiteSpace($clau)) {
            if ($null -eq $idxReq1 -or -not $idxReq1.ContainsKey($clau)) {
                [void]$orfes.Add($clau)
                continue
            }
            # El text mana a REQ1: aqui nomes se'n fa servir el cos.
            $cos = @($idxReq1[$clau].BodyLines)
        } else {
            $cos = @(_LlicCos $it)
        }
        $nod = _LlicFill $it 'nodisposa'
        $sid = _LlicFill $it 'sidisposa'
        $qua = _LlicFill $it 'quan'
        [void]$punts.Add([pscustomobject]@{
            Clau      = $clau
            Titol     = [string]$it.titol
            Condicio  = [string]$it.condicio
            Cos       = $cos
            NoDisposa = if ($null -ne $nod) { @(_LlicCos $nod) } else { @() }
            SiDisposa = if ($null -ne $sid) { @(_LlicCos $sid) } else { @() }
            Quan      = if ($null -ne $qua) { @(_LlicCos $qua) } else { @() }
            Subs      = @(@(_LlicFills $it 'subitem') | ForEach-Object { @(_LlicCos $_) })
        })
    }
    return @{ Punts = $punts.ToArray(); Orfes = $orfes.ToArray() }
}

# Frases que TANQUEN el bloc DESPRES d'un informe de Llicencia ja emes. Es
# comparen en majuscules i sense accents no: n'hi ha prou amb el principi.
function _LlicFinalsDeBloc {
    return @(
        'CONDICIONS LLIC',
        'ANNEX 1',
        'Ho poso al seu coneixement',
        'Cornell',
        "S'informa favorablement",
        [char]0x2018 + 'informa favorablement',   # apostrof tipografic del Word
        'Cal requerir'
    )
}

# Treu els punts del bloc DESPRES d'un informe de Llicencia JA EMES (el
# pre-llicencia), a partir del text dels seus paragrafs. Funcio PURA.
#
# PER QUE: el favorable POST-llicencia diu "Despres d'haver comprovat la seguent
# documentacio presentada:" i llista EXACTAMENT el que deia el pre-llicencia. Fer
# que l'usuari ho tornes a triar del cataleg era demanar-li que repetis una
# feina que ja consta escrita, i obria la porta a que les dues llistes no
# quadressin.
#
# COM ES RECONEIX: els informes els genera aquest mateix programa, i alli el
# numero i el pic s'escriuen com a TEXT (Format-Item escriu "N. ", Format-Bullet
# escriu U+2022 + tabulador); no hi ha numeracio automatica del Word. Per tant
# n'hi ha prou amb el text de cada paragraf.
#   "N. ..."  -> comenca un punt nou
#   U+2022    -> sub-punt del punt actual
#   "Quan: ..." -> ES DESCARTA (al post ja no toca: la documentacio ja s'ha
#                  presentat, o sigui que el termini no hi pinta res)
#   la resta  -> linia de cos del punt actual
function _LlicPuntsDelDocxAnterior($paraTexts) {
    $punts = New-Object System.Collections.ArrayList
    $dins = $false
    $actual = $null
    $pic = [string][char]0x2022

    foreach ($raw in @($paraTexts)) {
        $t = ([string]$raw).Trim()
        if ($t.Length -eq 0) { continue }

        if (-not $dins) {
            # El titol del bloc DESPRES. Es mira tolerant (l'usuari pot haver
            # retocat l'informe a ma) i en majuscules, que es com surt.
            $u = $t.ToUpper()
            if ($u.StartsWith('DOCUMENTACI') -and $u.Contains('DESPR')) { $dins = $true }
            continue
        }

        foreach ($fi in (_LlicFinalsDeBloc)) {
            if ($t.StartsWith($fi)) { $dins = $false; break }
        }
        if (-not $dins) { break }

        if ($t -match '^(\d+)\.\s+(.*)$') {
            if ($null -ne $actual) { [void]$punts.Add($actual) }
            $actual = [pscustomobject]@{
                Clau = ''; Titol = [string]$Matches[2]; Condicio = ''
                Cos = @([string]$Matches[2]); NoDisposa = @(); SiDisposa = @()
                Quan = @(); Subs = @()
            }
            continue
        }
        if ($null -eq $actual) { continue }   # text solt abans del primer punt

        if ($t.StartsWith($pic)) {
            $sub = $t.Substring(1).TrimStart([char]0x0009, ' ')
            $actual.Subs = @($actual.Subs) + @(, @($sub))
            continue
        }
        if ($t.StartsWith('Quan:')) { continue }
        $actual.Cos = @($actual.Cos) + @($t)
    }
    if ($null -ne $actual) { [void]$punts.Add($actual) }
    return $punts.ToArray()
}

# Quins punts condicionals entren, segons si es llicencia provisional.
#   'annexii'     -> nomes si NO ho es
#   'provisional' -> nomes si SI ho es
# Un punt sense condicio entra sempre. Funcio PURA.
function _LlicCondicioEntra([string]$condicio, [bool]$esProvisional) {
    $c = ([string]$condicio).Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($c)) { return $true }
    if ($c -eq 'provisional') { return $esProvisional }
    if ($c -eq 'annexii') { return (-not $esProvisional) }
    return $true
}

# Text de la conclusio d'una fase.
#   - Al favorable PRE, la coda " i sota les seguents condicions" es OPCIONAL:
#     no sempre n'hi ha. Sense condicions, la frase acaba amb un punt.
#   - Al favorable POST i si es llicencia provisional, s'hi afegeix la visita
#     d'inspeccio.
# Funcio PURA.
function _LlicConclusioText([string]$fase, [bool]$ambCondicions) {
    $f = @(_LlicFases) | Where-Object { $_.Clau -eq $fase } | Select-Object -First 1
    if ($null -eq $f) { return '' }
    if ($fase -ne 'favorable-pre') { return [string]$f.Conclusio }
    if ($ambCondicions) {
        return ([string]$f.Conclusio + ' i sota les seg' + [char]0x00FC + 'ents condicions.')
    }
    return ([string]$f.Conclusio + '.')
}

# Frase d'entrada del favorable POST-llicencia.
function _LlicEntradaPost([bool]$esProvisional) {
    $t = ('Despr' + [char]0x00E9 + 's d' + [char]0x2019 + 'haver comprovat la seg' + [char]0x00FC +
          'ent documentaci' + [char]0x00F3 + ' presentada')
    if ($esProvisional) {
        $t += (' i d' + [char]0x2019 + 'haver realitzat la posterior visita d' + [char]0x2019 + 'inspecci' + [char]0x00F3)
    }
    return ($t + ':')
}

# El paragraf "Documentacio signada digitalment pel tecnic redactor..." Funcio
# PURA. $data ja ve formatada com la vol l'usuari ("20 de febrer de 2024").
function _LlicTextDocumentacio([string]$tecnic, [string]$numCol, [string]$collegi, [string]$data) {
    if ([string]::IsNullOrWhiteSpace($tecnic)) { return '' }
    $t = ('Documentaci' + [char]0x00F3 + ' signada digitalment pel t' + [char]0x00E8 + 'cnic redactor ' + $tecnic.Trim())
    if (-not [string]::IsNullOrWhiteSpace($numCol)) {
        $t += (', col' + [char]0x00B7 + 'legiat n' + [char]0x00FA + 'mero ' + $numCol.Trim())
    }
    if (-not [string]::IsNullOrWhiteSpace($collegi)) { $t += (' del ' + $collegi.Trim()) }
    if (-not [string]::IsNullOrWhiteSpace($data)) { $t += (', en data ' + $data.Trim()) }
    return ($t + '.')
}

# Nom del fitxer de sortida. Segueix el mateix patro que la resta d'informes:
# data al principi (aixi "Actualitzar base d'informes" el reconeix).
function _LlicNomFitxer([datetime]$data, [string]$fase, [string]$idGia, [string]$titular) {
    $curt = switch ($fase) {
        'favorable-pre'  { 'LlicFavPre' }
        'favorable-post' { 'LlicFavPost' }
        default          { 'LlicReq' }
    }
    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add($data.ToString('yyyy-MM-dd'))
    [void]$parts.Add($curt)
    if (-not [string]::IsNullOrWhiteSpace($idGia)) { [void]$parts.Add('GIA ' + $idGia.Trim()) }
    if (-not [string]::IsNullOrWhiteSpace($titular)) { [void]$parts.Add($titular.Trim()) }
    $nom = ($parts -join '_')
    # Fora els caracters que Windows no admet en un nom de fitxer.
    $nom = [regex]::Replace($nom, '[\\/:*?"<>|]', '-')
    return ($nom + '.docx')
}

# ----------------------------------------------------------------------------
# COMPOSICIO DEL DOCUMENT (Word COM)
# ----------------------------------------------------------------------------
# Escriu un punt de Llicencia: el cos (de REQ1 o propi) com un item numerat, els
# seus sub-punts amb pic, i despres el comentari triat en VERD -negreta si falta
# la documentacio, sense negreta si ja hi es- i, al bloc DESPRES, el "Quan:".
#
# Tot el format surt de Format.ps1: aqui no s'hi inventa res. L'unic afegit es
# el color, que Format-Body ja sap aplicar.
function _LlicEscriuPunt($sel, $punt, [int]$numero, $fields, [string]$estat, [bool]$ambQuan) {
    $linies = @($punt.Cos)
    $primera = if ($linies.Count -gt 0) { [string]$linies[0] } else { [string]$punt.Titol }
    Format-Item $sel $numero (Apply-Fields -text $primera -fields $fields)
    for ($i = 1; $i -lt $linies.Count; $i++) {
        $t = [string]$linies[$i]
        if (_EsUrl $t) { Format-Url $sel (_TextUrl $t) } else { Format-Body $sel (Apply-Fields -text $t -fields $fields) }
    }
    # Sub-punts (per exemple, quines instal·lacions s'han de legalitzar).
    $primerSub = $true
    foreach ($sub in @($punt.Subs)) {
        foreach ($l in @($sub)) {
            if (_EsUrl $l) { Format-Url $sel (_TextUrl $l) -IsChild; continue }
            if ($primerSub) { Format-Bullet $sel (Apply-Fields -text $l -fields $fields) -First; $primerSub = $false }
            else { Format-Bullet $sel (Apply-Fields -text $l -fields $fields) }
        }
    }
    # El comentari. 'no' = falta (negreta); 'si' = ja hi es (normal). Al Word de
    # l'usuari anaven en verd, pero aquell color era una MARCA SEVA per saber que
    # havia de canviar a cada informe, no part del document: aqui van amb el
    # color de sempre (Format.ps1).
    $com = if ($estat -eq 'si') { @($punt.SiDisposa) } elseif ($estat -eq 'no') { @($punt.NoDisposa) } else { @() }
    $primerCom = $true
    foreach ($l in $com) {
        if (_EsUrl $l) { Format-Url $sel (_TextUrl $l); continue }
        $txt = Apply-Fields -text $l -fields $fields
        if ($primerCom -and $estat -eq 'no') { Format-Body $sel $txt -Bold }
        else { Format-Body $sel $txt }
        $primerCom = $false
    }
    if ($ambQuan) {
        foreach ($l in @($punt.Quan)) {
            Format-Body $sel ('Quan: ' + (Apply-Fields -text $l -fields $fields))
        }
    }
}

# Una linia del cos es un enllac? El motor marca els URL amb "[[URL]] ..." (el
# mateix conveni de tot el programa).
function _EsUrl([string]$linia) { return (([string]$linia).TrimStart() -like '[[URL]]*') }
function _TextUrl([string]$linia) {
    $t = ([string]$linia).TrimStart()
    if ($t -like '[[URL]]*') { return $t.Substring(7).Trim() }
    return $t
}

# Composa l'informe sencer i el desa. Retorna la ruta.
#
# $model porta tot el que ha triat l'usuari a l'assistent:
#   Fase, EsProvisional, Header, Fields, Abans, Projecte, Despres, Doc,
#   Condicions, Orfes.
function Build-LlicenciaDocument($word, $model) {
    $header = $model.Header
    $baseName = _LlicNomFitxer (Get-Date) ([string]$model.Fase) ([string]$header['ID_GIA']) ([string]$header['TITULAR'])
    $targetDir = _ResolveOutputDir
    [string]$outPath = _GetUniqueOutputPath $targetDir $baseName
    $fileName = [System.IO.Path]::GetFileName($outPath)
    $tempPath = Join-Path $env:TEMP $fileName
    $doc = _OpenOutputDocument $word $tempPath

    # La capcalera de LLICENCIA (porta la linia "Classificacio:"). Si el bloc no
    # hi es (0 CAPCALERA.docx encara sense actualitzar), Select-CapcaleraBlock es
    # queda amb el generic i l'informe surt igualment, sense la classificacio.
    Select-CapcaleraBlock $doc 'LLIC'
    Apply-HeaderReplacements -doc $doc -header $header

    $doc.Activate()
    $sel = $word.Selection
    [void]$sel.EndKey(6)   # wdStory

    $cfg = $Script:ReportFormatConfig
    $fields = $model.Fields
    $esPost = ([string]$model.Fase -eq 'favorable-post')

    if ($esPost) {
        # El post-llicencia no repeteix tot l'informe: nomes la documentacio que
        # s'ha comprovat, sense el "Quan:".
        Format-Body $sel (_LlicEntradaPost ([bool]$model.EsProvisional))
        $n = 0
        foreach ($p in @($model.Despres)) {
            $n++
            _LlicEscriuPunt $sel $p $n $fields '' $false
        }
    } else {
        # ---- ABANS ----
        Format-Section $sel (_LlicTitolAbans)
        $n = 0
        foreach ($p in @($model.Abans)) {
            $n++
            _LlicEscriuPunt $sel $p $n $fields ([string]$p.Estat) $false
        }
        # ---- PROJECTE: els requeriments normals de REQ1, com sempre ----
        $proj = @($model.Projecte)
        if ($proj.Count -gt 0) {
            Format-Spacer $sel
            Format-Subsection $sel 'Projecte'
            foreach ($p in $proj) {
                $n++
                _LlicEscriuPunt $sel $p $n $fields '' $false
            }
        }
        # ---- DOCUMENTACIO del tecnic redactor ----
        $doc1 = [string]$model.Doc.Text
        if (-not [string]::IsNullOrWhiteSpace($doc1)) {
            Format-Spacer $sel
            Format-Subsection $sel ('Documentaci' + [char]0x00F3)
            $n++
            Format-Item $sel $n $doc1
            $primer = $true
            foreach ($d in @($model.Doc.Items)) {
                if ($primer) { Format-Bullet $sel ([string]$d) -First; $primer = $false }
                else { Format-Bullet $sel ([string]$d) }
            }
        }
        # ---- DESPRES ----
        $desp = @($model.Despres)
        if ($desp.Count -gt 0) {
            Format-Spacer $sel
            Format-Section $sel (_LlicTitolDespres)
            $n = 0
            foreach ($p in $desp) {
                $n++
                _LlicEscriuPunt $sel $p $n $fields ([string]$p.Estat) $true
            }
        }
    }

    # ---- CONCLUSIO ----
    $ambCond = (-not [string]::IsNullOrWhiteSpace([string]$model.Condicions))
    Format-Spacer $sel
    Format-Conclusion $sel (_LlicConclusioText ([string]$model.Fase) $ambCond)
    if ($ambCond -and [string]$model.Fase -eq 'favorable-pre') {
        Format-Spacer $sel
        Format-Section $sel ('CONDICIONS LLIC' + [char]0x00C8 + 'NCIA')
        foreach ($l in (([string]$model.Condicions) -split "`r?`n")) {
            if ([string]::IsNullOrWhiteSpace($l)) { continue }
            Format-Body $sel ([string]$l).Trim() -Bold
        }
    }
    # El tancament de sempre, igual que a la resta d'informes.
    Format-Spacer $sel
    Format-Body $sel ('Ho poso al seu coneixement als efectes oportuns,')
    Format-Body $sel ('Cornell' + [char]0x00E0 + ' de Llobregat,')

    # ---- ANNEX 1: nomes al REQUERIMENT d'una llicencia PROVISIONAL ----
    if ([string]$model.Fase -eq 'requeriment' -and [bool]$model.EsProvisional) {
        _LlicEscriuAnnex1 $sel $model.Cataleg
    }

    $doc.Save()
    $doc.Close($false)
    try { Move-Item -LiteralPath $tempPath -Destination $outPath -Force } catch { return $tempPath }
    return $outPath
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

function _LlicEscriuAnnex1($sel, $llic) {
    $sec = _LlicSeccioAnnex1 $llic
    if ($null -eq $sec) { return }
    # Salt de pagina: l'annex es un document a part dins de l'informe.
    try { [void]$sel.InsertBreak(7) } catch { Format-Spacer $sel }   # wdPageBreak
    Format-Section $sel ([string]$sec.titol)
    $n = 0
    $primerSub = $true
    foreach ($nd in @($sec.fills)) {
        $tipus = [string]$nd.tipus
        foreach ($l in @(_LlicCos $nd)) {
            if (_EsUrl $l) { Format-Url $sel (_TextUrl $l); continue }
            switch ($tipus) {
                'item'    { $n++; Format-Item $sel $n $l; $primerSub = $true }
                'subitem' {
                    if ($primerSub) { Format-Bullet $sel $l -First; $primerSub = $false }
                    else { Format-Bullet $sel $l }
                }
                default   { Format-Body $sel $l }
            }
        }
    }
}

# ----------------------------------------------------------------------------
# ASSISTENT (WinForms, nomes Windows)
# ----------------------------------------------------------------------------
# Pas 1: la FASE i si es llicencia provisional. Retorna @{ Nav; Fase; Prov }.
function Select-LlicFase($preFase, $preProv) {
    $form = _NewForm
    $form.Text = 'Llic' + [char]0x00E8 + 'ncia - Pas 1'
    $form.ClientSize = New-Object System.Drawing.Size(520, 330)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 72)
    $lbl.Size = New-Object System.Drawing.Size(480, 20)
    $lbl.Text = 'Quin dels tres informes de la llic' + [char]0x00E8 + 'ncia vols fer?'
    [void]$form.Controls.Add($lbl)

    $y = 98
    $radios = @{}
    foreach ($f in @(_LlicFases)) {
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
    if (-not ($radios.Values | Where-Object { $_.Checked })) { $radios['requeriment'].Checked = $true }

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

    $res = @{ Nav = 'back' }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Continuar'
    $btnOk.Location = New-Object System.Drawing.Point(370, 286)
    $btnOk.Size = New-Object System.Drawing.Size(130, 32)
    _StylePrimaryButton $btnOk
    $btnOk.add_Click({
        foreach ($k in $radios.Keys) { if ($radios[$k].Checked) { $res.Fase = $k } }
        $res.Prov = [bool]$cbProv.Checked
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())
    [void]$form.Controls.Add($btnOk)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = [string][char]0x2190 + ' Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(20, 286)
    $btnBack.Size = New-Object System.Drawing.Size(110, 32)
    _StyleSecondaryButton $btnBack
    $btnBack.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnBack)

    [void](_AddBrandHeader $form ('Llic' + [char]0x00E8 + 'ncia (Annex II / LL Prov)') 'Tria quin informe vols fer' 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

# Pas de tria de DOCUMENTACIO (blocs ABANS i DESPRES): per cada punt, si aplica
# i, si aplica, si la documentacio ja hi es o no.
# Retorna @{ Nav; Punts } amb els punts triats i el seu Estat ('no' | 'si').
# $marcatPerDefecte: si els punts surten ja marcats. Al bloc DESPRES si (el Word
# de l'usuari els portava tots i ell hi anava esborrant el que no tocava; picar
# quinze caselles cada vegada era feina de mes), i al bloc ABANS no, perque alli
# cada punt demana a mes decidir si es te la documentacio o no.
function Select-LlicDocumentacio($punts, [string]$titol, [string]$subtitol, [bool]$ambEstat, [bool]$marcatPerDefecte = $false) {
    $punts = @($punts)
    $form = _NewForm
    $form.Text = $titol
    $form.ClientSize = New-Object System.Drawing.Size(900, 620)
    $form.StartPosition = 'CenterScreen'

    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid
    # ATENCIO: _StyleListGrid fa Dock='Fill', pensat per a graelles que viuen
    # DINS d'un panell (Editar base, Controls periodics). Aqui la graella
    # conviu amb els botons posats a ma al mateix formulari: amb el Dock posat
    # ocupava TOTA la finestra i els TAPAVA -cap boto visible ni clicable, la
    # pantalla semblava morta-. Per aixo el Dock es desfa i la posicio va
    # DESPRES de l'estil (abans, l'estil la trepitjava).
    $grid.Dock = 'None'
    $grid.Location = New-Object System.Drawing.Point(15, 70)
    $grid.Size = New-Object System.Drawing.Size(870, 480)
    $grid.Anchor = 'Top,Bottom,Left,Right'
    $grid.AllowUserToAddRows = $false
    $grid.AutoGenerateColumns = $false

    $cAplica = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $cAplica.HeaderText = 'Aplica'; $cAplica.Width = 55
    [void]$grid.Columns.Add($cAplica)

    $cPunt = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cPunt.HeaderText = 'Punt'; $cPunt.Width = 560; $cPunt.ReadOnly = $true
    [void]$grid.Columns.Add($cPunt)

    if ($ambEstat) {
        $cEstat = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
        $cEstat.HeaderText = 'Documentaci' + [char]0x00F3
        $cEstat.Width = 220
        [void]$cEstat.Items.Add('No es disposa')
        [void]$cEstat.Items.Add('Es disposa')
        [void]$grid.Columns.Add($cEstat)
    }
    [void]$form.Controls.Add($grid)

    foreach ($p in $punts) {
        $txt = if (@($p.Cos).Count -gt 0) { [string]@($p.Cos)[0] } else { [string]$p.Titol }
        $i = $grid.Rows.Add()
        $grid.Rows[$i].Cells[0].Value = $marcatPerDefecte
        $grid.Rows[$i].Cells[1].Value = $txt
        if ($ambEstat) { $grid.Rows[$i].Cells[2].Value = 'No es disposa' }
        $grid.Rows[$i].Tag = $p
    }

    $res = @{ Nav = 'back'; Punts = @() }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Continuar'
    $btnOk.Location = New-Object System.Drawing.Point(760, 566)
    $btnOk.Size = New-Object System.Drawing.Size(125, 34)
    $btnOk.Anchor = 'Bottom,Right'
    _StylePrimaryButton $btnOk
    $btnOk.add_Click({
        $sel = New-Object System.Collections.ArrayList
        foreach ($row in $grid.Rows) {
            if (-not [bool]$row.Cells[0].Value) { continue }
            $p = $row.Tag
            $estat = if ($ambEstat -and [string]$row.Cells[2].Value -eq 'Es disposa') { 'si' } else { 'no' }
            [void]$sel.Add(([pscustomobject]@{
                Clau = $p.Clau; Titol = $p.Titol; Condicio = $p.Condicio
                Cos = $p.Cos; NoDisposa = $p.NoDisposa; SiDisposa = $p.SiDisposa
                Quan = $p.Quan; Subs = $p.Subs; Estat = $estat
            }))
        }
        $res.Punts = $sel.ToArray()
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())
    [void]$form.Controls.Add($btnOk)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = [string][char]0x2190 + ' Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(15, 566)
    $btnBack.Size = New-Object System.Drawing.Size(115, 34)
    $btnBack.Anchor = 'Bottom,Left'
    _StyleSecondaryButton $btnBack
    $btnBack.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnBack)

    # Marcar-ho / desmarcar-ho tot: amb quinze punts, anar picant casella a
    # casella es el que fa que l'eina no compensi.
    $marcaTot = {
        param($valor)
        foreach ($row in $grid.Rows) { $row.Cells[0].Value = $valor }
        $grid.EndEdit() | Out-Null
    }.GetNewClosure()
    $btnTot = New-Object System.Windows.Forms.Button
    $btnTot.Text = 'Marcar-ho tot'
    $btnTot.Location = New-Object System.Drawing.Point(140, 566)
    $btnTot.Size = New-Object System.Drawing.Size(125, 34)
    $btnTot.Anchor = 'Bottom,Left'
    _StyleSecondaryButton $btnTot
    $btnTot.add_Click({ & $marcaTot $true }.GetNewClosure())
    [void]$form.Controls.Add($btnTot)

    $btnCap = New-Object System.Windows.Forms.Button
    $btnCap.Text = 'Desmarcar-ho tot'
    $btnCap.Location = New-Object System.Drawing.Point(273, 566)
    $btnCap.Size = New-Object System.Drawing.Size(140, 34)
    $btnCap.Anchor = 'Bottom,Left'
    _StyleSecondaryButton $btnCap
    $btnCap.add_Click({ & $marcaTot $false }.GetNewClosure())
    [void]$form.Controls.Add($btnCap)

    [void](_AddBrandHeader $form $titol $subtitol 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

# Pas de la DOCUMENTACIO del tecnic redactor.
# Retorna @{ Nav; Text; Items }.
function Select-LlicTecnic($pre) {
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

    $docs = @('Projecte', 'Pl' + [char]0x00E0 + 'nols', 'Annexos')
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
        $cbDoc[[string]$d] = $cb
        $tbDoc[[string]$d] = $t
        $y += 30
    }

    $res = @{ Nav = 'back'; Text = ''; Items = @(); Camps = @{} }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Continuar'
    $btnOk.Location = New-Object System.Drawing.Point(465, 380)
    $btnOk.Size = New-Object System.Drawing.Size(130, 32)
    _StylePrimaryButton $btnOk
    $btnOk.add_Click({
        $res.Text = _LlicTextDocumentacio $tb['Tecnic'].Text $tb['NumCol'].Text $tb['Collegi'].Text $tb['Data'].Text
        $items = New-Object System.Collections.ArrayList
        foreach ($d in $docs) {
            if (-not $cbDoc[[string]$d].Checked) { continue }
            $id = ([string]$tbDoc[[string]$d].Text).Trim()
            if ([string]::IsNullOrWhiteSpace($id)) { [void]$items.Add([string]$d) }
            else { [void]$items.Add([string]$d + ' (Id Firmadoc: ' + $id + ')') }
        }
        $res.Items = $items.ToArray()
        foreach ($k in $tb.Keys) { $res.Camps[$k] = [string]$tb[$k].Text }
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())
    [void]$form.Controls.Add($btnOk)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = [string][char]0x2190 + ' Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(20, 380)
    $btnBack.Size = New-Object System.Drawing.Size(115, 32)
    _StyleSecondaryButton $btnBack
    $btnBack.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnBack)

    [void](_AddBrandHeader $form ('Documentaci' + [char]0x00F3) ('Qui ha signat el projecte i amb quin Id Firmadoc') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

# Pas de les CONDICIONS (nomes al favorable pre-llicencia): quadre de text
# lliure. Buit = la conclusio acaba sense "i sota les seguents condicions".
function Select-LlicCondicions([string]$pre) {
    $form = _NewForm
    $form.Text = 'Condicions de la llic' + [char]0x00E8 + 'ncia'
    $form.ClientSize = New-Object System.Drawing.Size(660, 440)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 70)
    $lbl.Size = New-Object System.Drawing.Size(620, 36)
    $lbl.Text = ('Una condici' + [char]0x00F3 + ' per l' + [char]0x00ED + 'nia. Si ho deixes BUIT, la conclusi' +
                 [char]0x00F3 + ' acaba a "...donar per tancat l' + [char]0x2019 + 'expedient." i no surt' + "`r`n" +
                 'el bloc CONDICIONS LLIC' + [char]0x00C8 + 'NCIA.')
    [void]$form.Controls.Add($lbl)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(20, 112)
    $tb.Size = New-Object System.Drawing.Size(620, 264)
    $tb.Multiline = $true
    $tb.ScrollBars = 'Vertical'
    $tb.Anchor = 'Top,Bottom,Left,Right'
    $tb.Text = [string]$pre
    [void]$form.Controls.Add($tb)

    $res = @{ Nav = 'back'; Text = '' }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Continuar'
    $btnOk.Location = New-Object System.Drawing.Point(510, 390)
    $btnOk.Size = New-Object System.Drawing.Size(130, 32)
    $btnOk.Anchor = 'Bottom,Right'
    _StylePrimaryButton $btnOk
    $btnOk.add_Click({ $res.Text = [string]$tb.Text; $res.Nav = 'fwd'; $form.DialogResult = 'OK'; $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnOk)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = [string][char]0x2190 + ' Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(20, 390)
    $btnBack.Size = New-Object System.Drawing.Size(115, 32)
    $btnBack.Anchor = 'Bottom,Left'
    _StyleSecondaryButton $btnBack
    $btnBack.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnBack)

    [void](_AddBrandHeader $form 'Condicions' ('Nom' + [char]0x00E9 + 's al favorable pre-llic' + [char]0x00E8 + 'ncia') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

# ----------------------------------------------------------------------------
# PUNT D'ENTRADA (des del menu)
# ----------------------------------------------------------------------------
function Invoke-LlicenciaWizard {
    $llic = Read-LlicCataleg
    if ($null -eq $llic) {
        [System.Windows.Forms.MessageBox]::Show(
            ("No s'ha trobat ESTRUCTURALS\LLIC.json.`n`nAquest fitxer es la base de dades de Llicencia " +
             "(que aporta cada requeriment i que no). Fes 'Actualitzar.bat' per baixar-lo."),
            'Llicencia', 'OK', 'Error') | Out-Null
        return
    }

    $word = $null
    # $st.Fields es el diccionari de camps COMPARTIT de tot l'assistent (el
    # mateix paper que a Invoke-NouWizard): els [CAMP:]/[OPCIO:] s'hi omplen
    # alla on surten i despres la composicio els hi busca.
    $st = @{ Fase = 'requeriment'; Prov = $false; Tecnic = @{}; Condicions = ''
             Fields = [ordered]@{}; PostLlegit = $false }
    $step = 1
    try {
        while ($true) {
            switch ($step) {
                1 {
                    $r = Select-LlicFase $st.Fase $st.Prov
                    if ($r.Nav -ne 'fwd') { return }
                    $st.Fase = [string]$r.Fase
                    $st.Prov = [bool]$r.Prov
                    $step = 2
                }
                2 {
                    $r = Get-HeaderData -preload $st.HeaderPre
                    if ($r.Nav -eq 'back') { $step = 1; break }
                    $st.Header = $r.Data
                    $st.HeaderPre = $r.Data
                    $step = 3
                }
                3 {
                    # Aqui nomes cal REQ1 (el JSON d'on surt el text). El Word
                    # s'arrenca DIFERIT al pas 9, quan es genera de debo (mateix
                    # motiu que a Invoke-NouWizard: arrencar-lo en fred es lent
                    # i si l'usuari tira enrere no ha de quedar obert per res).
                    if ($null -eq $st.Req1) {
                        $req1Path = Join-Path $EstructuralsDir 'REQ1.json'
                        $st.Req1 = Get-ParsedCataleg -path $req1Path
                        $st.IdxReq1 = _LlicIndexReq1 $st.Req1
                    }
                    # Els punts, resolts amb el text de REQ1. Les claus ORFES
                    # s'avisen: si algu ha reanomenat un requeriment a REQ1, el
                    # punt desapareixeria de l'informe sense dir res.
                    $bAbans  = _LlicPuntsPerBloc $llic $st.IdxReq1 'ABANS'
                    $bDesp   = _LlicPuntsPerBloc $llic $st.IdxReq1 'DESPRES'
                    $bPropis = _LlicPuntsPerBloc $llic $st.IdxReq1 'PROPIS'
                    $orfes = @($bAbans.Orfes) + @($bDesp.Orfes) + @($bPropis.Orfes)
                    if ($orfes.Count -gt 0) {
                        [System.Windows.Forms.MessageBox]::Show(
                            ("Hi ha " + $orfes.Count + " punt(s) de Llicencia que apunten a un requeriment que JA NO " +
                             "existeix a REQ1:`n`n  " + ($orfes -join "`n  ") +
                             "`n`nNo sortiran a l'informe. Arregla-ho des de l'editor de catalegs."),
                            'Llicencia', 'OK', 'Warning') | Out-Null
                    }
                    # Els condicionals entren segons el tipus de llicencia.
                    $cond = @(@($bPropis.Punts) | Where-Object { _LlicCondicioEntra ([string]$_.Condicio) ([bool]$st.Prov) })
                    $st.AbansTots = @($cond) + @($bAbans.Punts)
                    $st.DespresTots = @($bDesp.Punts)
                    $step = 4
                }
                4 {
                    if ([string]$st.Fase -eq 'favorable-post') {
                        # El POST no repeteix la tria: LLEGEIX el pre-llicencia i
                        # en treu la documentacio que hi constava. Si no se'n pot
                        # treure res, es cau a la llista del cataleg (l'informe
                        # s'ha de poder fer igualment).
                        if (-not $st.PostLlegit) {
                            $anterior = Select-PreviousReport
                            if (-not $anterior) { $step = 2; break }
                            $st.PostLlegit = $true
                            $delDoc = @()
                            try {
                                $xmlInfo = _LoadDocxXml $anterior
                                $paras = @(_BodyParagraphsXml $xmlInfo)
                                $delDoc = @(_LlicPuntsDelDocxAnterior (@($paras | ForEach-Object { _ParagraphTextXml $_ $xmlInfo.Ns })))
                            } catch {
                                $delDoc = @()
                            }
                            if ($delDoc.Count -eq 0) {
                                [System.Windows.Forms.MessageBox]::Show(
                                    ("D'aquest informe no n'he pogut treure el bloc " +
                                     [char]0x2018 + 'DOCUMENTACI' + [char]0x00D3 + ' NECESS' + [char]0x00C0 + 'RIA DESPR' + [char]0x00C9 + 'S...' + [char]0x2019 +
                                     ".`n`nSegurament no es un informe de Llicencia (o s'ha retocat molt a ma). " +
                                     "Et deixo la llista del cataleg per triar-la a ma."),
                                    'Llicencia', 'OK', 'Warning') | Out-Null
                            } else {
                                $st.DespresTots = $delDoc
                            }
                        }
                        $step = 7
                        break
                    }
                    $r = Select-LlicDocumentacio $st.AbansTots ('Documentaci' + [char]0x00F3 + ' ABANS de la resoluci' + [char]0x00F3) `
                            ('Marca la que aplica i si ja es t' + [char]0x00E9) $true $false
                    if ($r.Nav -ne 'fwd') { $step = 2; break }
                    $st.Abans = $r.Punts
                    $step = 5
                }
                5 {
                    # Projecte: els requeriments NORMALS de REQ1, amb la mateixa
                    # pantalla de sempre (no s'hi inventa res).
                    $r = Select-Items -sections $st.Req1.Sections -preloadSelectedKeys $st.ProjKeys -fields $st.Fields -preloadValues $st.ProjVals
                    if ($r.Nav -eq 'back') { $step = 4; break }
                    if ($r.Nav -eq 'stay') { break }
                    $st.ProjSel = $r.Data
                    $st.ProjKeys = Get-SelectedKeysFromResult $st.ProjSel
                    $st.ProjVals = Get-FieldValuesForSession $st.Fields
                    $step = 6
                }
                6 {
                    $r = Select-LlicTecnic $st.Tecnic
                    if ($r.Nav -ne 'fwd') { $step = 5; break }
                    $st.Doc = @{ Text = [string]$r.Text; Items = @($r.Items) }
                    $st.Tecnic = $r.Camps
                    $step = 7
                }
                7 {
                    $esPost = ([string]$st.Fase -eq 'favorable-post')
                    $titol = if ($esPost) {
                        'Documentaci' + [char]0x00F3 + ' comprovada'
                    } else {
                        'Documentaci' + [char]0x00F3 + ' DESPR' + [char]0x00C9 + 'S de la resoluci' + [char]0x00F3
                    }
                    $sub = if ($esPost) { ('Ve de l' + [char]0x2019 + 'informe anterior; treu el que no s' + [char]0x2019 + 'hagi comprovat') }
                           else         { "Marca la que entra a l'informe" }
                    # Tot marcat de sortida: al POST ve de l'informe anterior (hi
                    # ha de constar tot) i al pre-llicencia el Word de l'usuari
                    # tambe els portava tots.
                    $r = Select-LlicDocumentacio $st.DespresTots $titol $sub (-not $esPost) $true
                    if ($r.Nav -ne 'fwd') {
                        # Enrere al POST = tornar a triar l'informe anterior, que
                        # es l'unica cosa que hi ha darrere.
                        if ($esPost) { $st.PostLlegit = $false; $step = 4 } else { $step = 6 }
                        break
                    }
                    $st.Despres = $r.Punts
                    $step = 8
                }
                8 {
                    if ([string]$st.Fase -ne 'favorable-pre') { $step = 9; break }
                    $r = Select-LlicCondicions $st.Condicions
                    if ($r.Nav -ne 'fwd') { $step = 7; break }
                    $st.Condicions = [string]$r.Text
                    $step = 9
                }
                9 {
                    if ($null -eq $word) { $word = New-WordApp }
                    # Els camps [CAMP: ...] dels textos triats.
                    $model = @{
                        Fase = [string]$st.Fase
                        EsProvisional = [bool]$st.Prov
                        Header = $st.Header
                        Fields = $st.Fields
                        Abans = @($st.Abans)
                        Projecte = @(_LlicPuntsDeSeleccio $st.ProjSel)
                        Despres = @($st.Despres)
                        Doc = $(if ($null -ne $st.Doc) { $st.Doc } else { @{ Text = ''; Items = @() } })
                        Condicions = [string]$st.Condicions
                        Cataleg = $llic
                    }
                    $out = Build-LlicenciaDocument $word $model
                    # Es desa igual que a "Requeriment - Nou" perque el boto
                    # "Recuperar dades ultim informe" del Pas 2 hi arribi. Les
                    # claus desades son les de REQ1 (el bloc Projecte es el
                    # mateix cataleg), o sigui que serveixen als dos fluxos.
                    Save-LastReport ([ordered]@{
                        Version         = 1
                        Timestamp       = (Get-Date).ToString('o')
                        CatalegBaseName = 'LLIC'
                        Header          = $st.Header
                        SelectedKeys    = @($st.ProjKeys)
                        FieldValues     = (Get-FieldValuesForSession $st.Fields)
                        ConclusionTexts = @()
                    })
                    [System.Windows.Forms.MessageBox]::Show(
                        "Informe generat:`n$out", 'Finalitzat', 'OK', 'Information') | Out-Null
                    # Es deixa el Word obert amb l'informe, com a la resta de fluxos.
                    $word.Visible = $true
                    $word.Documents.Open($out) | Out-Null
                    return
                }
                default { return }
            }
        }
    } finally {
        # Si l'informe s'ha generat, el Word s'ha fet visible i es deixa obert
        # per a l'usuari; si es va cancel·lar pel cami, es tanca.
        if ($null -ne $word -and -not $word.Visible) { try { Close-WordApp $word } catch { } }
    }
}

# Converteix el que retorna Select-Items (les deficiencies normals de REQ1) als
# mateixos punts que fa servir la composicio, per no tenir dos camins.
function _LlicPuntsDeSeleccio($seleccio) {
    $out = New-Object System.Collections.ArrayList
    foreach ($sec in @($seleccio)) {
        foreach ($it in @($sec.Items)) {
            $subs = New-Object System.Collections.ArrayList
            foreach ($ch in @($it.Children)) { [void]$subs.Add(@($ch.BodyLines)) }
            [void]$out.Add([pscustomobject]@{
                Clau = ''; Titol = [string]$it.Short; Condicio = ''
                Cos = @($it.BodyLines); NoDisposa = @(); SiDisposa = @()
                Quan = @(); Subs = $subs.ToArray(); Estat = ''
            })
        }
    }
    return $out.ToArray()
}
