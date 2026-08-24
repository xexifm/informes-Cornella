#requires -Version 5.1
<#
.SYNOPSIS
  La CAPCALERA dels informes, en JSON i editable des del programa.

.DESCRIPTION
  '0 CAPCALERA.docx' es l'UNICA plantilla de Word de veritat que queda: hi ha
  l'escut, la taula del requadre de la "Nota:", les parades de tabulacio i els
  marges, i no es pot regenerar des de cap model. Per aixo NO es converteix a
  JSON: el que es fa es separar les dues coses.

    EL .docx mana en el FORMAT   (escut, taula, tipus de lletra, tabulacions)
    EL .json mana en el TEXT     (etiquetes, valors fixos, la nota)

  '0 CAPCALERA.json' es genera del propi .docx (_CapBlocsDelXml) i, en desar-lo
  des de l'editor de catalegs, el text que hi hagi es torna a escriure DINS del
  .docx (_CapAplicaAlXml). Aixi l'usuari pot canviar "Exp. Num:" o el text de
  la nota des del programa, i no es toca res mes del document.

  ATENCIO: '0 CAPCALERA.docx' NO ES POT SERIALITZAR AMB UN SERIALITZADOR D'XML.
  Ja va passar (vegeu CLAUDE.md): dels 19 espais de noms de l'arrel en van
  quedar 3, el Word va dir que el fitxer estava corromput i no es va poder
  generar CAP informe durant dies. Per aixo aqui tot son EDICIONS DE TEXT sobre
  'word/document.xml' -es busquen els trossos i s'hi empalma- i el ZIP es
  reescriu conservant l'ordre i la compressio de cada entrada.

  ELS BLOCS. El document en porta tres, un darrere l'altre, separats per un
  paragraf amb el marcador [[CAP:X]]:
    (cap marcador) -> el generic  : Requeriment - Nou, Ampliacio de termini...
    [[CAP:ACT_EXTR]]              : activitats extraordinaries
    [[CAP:LLIC]]                  : llicencia (i els seus informes curts)
  Select-CapcaleraBlock (Document.ps1) es queda amb el que toca i esborra els
  altres.

.NOTES
  Les funcions de text son PURES i es proven a Linux sobre el .docx real.
  CONVENCIO ASCII: el codi no porta accents.
#>

function Get-CapcaleraDocxPath { return [string](Join-Path $EstructuralsDir '0 CAPCALERA.docx') }
function Get-CapcaleraJsonPath { return [string](Join-Path $EstructuralsDir '0 CAPCALERA.json') }

# A QUIN TIPUS D'INFORME S'APLICA cada bloc. Funcio PURA. Va al JSON perque
# l'editor ho pugui ensenyar: "quin document estic tocant" era justament el que
# no quedava clar.
function _CapAplicaDe([string]$clau) {
    switch (([string]$clau).Trim().ToUpper()) {
        # ATENCIO als PARENTESIS: dins d'un @(...) la coma lliga MES FORT que el
        # '+' (vegeu CLAUDE.md). Sense ells, "Ampliaci' + [char]0xF3 + ' de..."
        # son TRES elements de la llista i el text surt esmicolat.
        'ACT_EXTR' { return @(('Activitats extraordin' + [char]0x00E0 + 'ries')) }
        'LLIC'     { return @(('Llic' + [char]0x00E8 + 'ncia (Annex II / LL Prov)'),
                              ('Modificaci' + [char]0x00F3 + ' NO Substancial'),
                              ('Trasp' + [char]0x00E0 + 's')) }
    }
    return @('Requeriment - Nou (REQ1)',
             ('Ampliaci' + [char]0x00F3 + ' de termini (TERMINI)'),
             ('Controls peri' + [char]0x00F2 + 'dics'))
}

function _CapTitolDe([string]$clau) {
    switch (([string]$clau).Trim().ToUpper()) {
        'ACT_EXTR' { return 'Capcalera - activitats extraordinaries' }
        'LLIC'     { return 'Capcalera - llicencia' }
    }
    return 'Capcalera - general'
}

# ----------------------------------------------------------------------------
# LECTURA DEL XML (tot per TEXT, mai amb un serialitzador)
# ----------------------------------------------------------------------------
# Els <w:p> del cos, amb on comencen i on acaben dins de la cadena. Funcio PURA.
function _CapTrossejaParagrafs([string]$xml) {
    $out = New-Object System.Collections.ArrayList
    $rx = [regex]'<w:p(?:\s[^>]*)?(?:/>|>)'
    $pos = 0
    while ($true) {
        $m = $rx.Match($xml, $pos)
        if (-not $m.Success) { break }
        if ($m.Value.EndsWith('/>')) {
            [void]$out.Add(@{ Ini = $m.Index; Fi = ($m.Index + $m.Length); Xml = $m.Value })
            $pos = $m.Index + $m.Length
            continue
        }
        $fi = $xml.IndexOf('</w:p>', $m.Index)
        if ($fi -lt 0) { break }
        $fi += 6
        [void]$out.Add(@{ Ini = $m.Index; Fi = $fi; Xml = $xml.Substring($m.Index, $fi - $m.Index) })
        $pos = $fi
    }
    return $out.ToArray()
}

# Els <w:r> d'un paragraf, amb el seu text, si van en negreta i si son un
# tabulador. Els offsets son RELATIUS al paragraf. Funcio PURA.
function _CapRunsDeParagraf([string]$pXml) {
    $out = New-Object System.Collections.ArrayList
    $rx = [regex]'<w:r(?:\s[^>]*)?>'
    $pos = 0
    while ($true) {
        $m = $rx.Match($pXml, $pos)
        if (-not $m.Success) { break }
        $fi = $pXml.IndexOf('</w:r>', $m.Index)
        if ($fi -lt 0) { break }
        $fi += 6
        $run = $pXml.Substring($m.Index, $fi - $m.Index)
        # La negreta es <w:b/> o <w:b w:val="..."/>; el <w:bCs/> es d'un altre
        # alfabet i no compta.
        $negreta = [bool]([regex]::IsMatch($run, '<w:b(?:\s[^>]*)?/>|<w:b(?:\s[^>]*)?>'))
        $esTab = $run.Contains('<w:tab/>')
        $txt = ''
        foreach ($t in [regex]::Matches($run, '<w:t(?:\s[^>]*)?>(.*?)</w:t>', 'Singleline')) {
            $txt += [string]$t.Groups[1].Value
        }
        [void]$out.Add(@{
            Ini = $m.Index; Fi = $fi; Xml = $run
            Text = (_CapDesescapa $txt); Negreta = $negreta; EsTab = $esTab
        })
        $pos = $fi
    }
    return $out.ToArray()
}

function _CapDesescapa([string]$t) {
    return ([string]$t).Replace('&lt;', '<').Replace('&gt;', '>').Replace('&quot;', '"').Replace('&apos;', "'").Replace('&amp;', '&')
}

function _CapEscapa([string]$t) {
    return ([string]$t).Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

# El text sencer d'un paragraf (per reconeixer els marcadors [[CAP:X]]).
function _CapTextDeParagraf([string]$pXml) {
    $t = ''
    foreach ($m in [regex]::Matches($pXml, '<w:t(?:\s[^>]*)?>(.*?)</w:t>', 'Singleline')) {
        $t += [string]$m.Groups[1].Value
    }
    return (_CapDesescapa $t)
}

# UNA LINIA de la capcalera a partir dels seus runs. Funcio PURA.
#
#   etiqueta -> "ID GIA:" (negreta) + tabulador + "<<ID_GIA>>" (normal)
#   text     -> un paragraf sense etiqueta ("INFORME", la nota, el departament)
#   buida    -> un paragraf sense text (no s'ensenya ni es toca)
function _CapLiniaDeRuns($runs) {
    $rs = @($runs)
    $iTab = -1
    for ($i = 0; $i -lt $rs.Count; $i++) { if ($rs[$i].EsTab) { $iTab = $i; break } }
    $etiqueta = ''
    $valor = ''
    if ($iTab -ge 0) {
        for ($i = 0; $i -lt $iTab; $i++) { $etiqueta += [string]$rs[$i].Text }
        for ($i = $iTab + 1; $i -lt $rs.Count; $i++) { $valor += [string]$rs[$i].Text }
    } else {
        # Sense tabulador: el que va en negreta al principi es l'etiqueta i la
        # resta el valor ("Classificacio: <<CLASSIFICACIO>>", "Nota: ...").
        $i = 0
        while ($i -lt $rs.Count -and [bool]$rs[$i].Negreta) { $etiqueta += [string]$rs[$i].Text; $i++ }
        for (; $i -lt $rs.Count; $i++) { $valor += [string]$rs[$i].Text }
    }
    $tot = ($etiqueta + $valor)
    if ([string]::IsNullOrWhiteSpace($tot)) { return @{ Tipus = 'buida'; Etiqueta = ''; Valor = '' } }
    if ([string]::IsNullOrWhiteSpace($valor)) { return @{ Tipus = 'text'; Etiqueta = ''; Valor = $etiqueta } }
    if ([string]::IsNullOrWhiteSpace($etiqueta)) { return @{ Tipus = 'text'; Etiqueta = ''; Valor = $valor } }
    return @{ Tipus = 'etiqueta'; Etiqueta = $etiqueta; Valor = $valor }
}

# ELS TRES BLOCS del document, amb les seves linies. Funcio PURA.
function _CapBlocsDelXml([string]$xml) {
    $paras = @(_CapTrossejaParagrafs $xml)
    $blocs = New-Object System.Collections.ArrayList
    $actual = @{ Clau = ''; Linies = (New-Object System.Collections.ArrayList) }
    [void]$blocs.Add($actual)
    for ($i = 0; $i -lt $paras.Count; $i++) {
        $txt = (_CapTextDeParagraf ([string]$paras[$i].Xml)).Trim()
        $m = [regex]::Match($txt, '^\[\[CAP:([^\]]*)\]\]$')
        if ($m.Success) {
            $actual = @{ Clau = [string]$m.Groups[1].Value; Linies = (New-Object System.Collections.ArrayList) }
            [void]$blocs.Add($actual)
            continue
        }
        $l = _CapLiniaDeRuns (_CapRunsDeParagraf ([string]$paras[$i].Xml))
        $l['Para'] = $i
        [void]$actual.Linies.Add($l)
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($b in $blocs) { [void]$out.Add(@{ Clau = [string]$b.Clau; Linies = $b.Linies.ToArray() }) }
    return $out.ToArray()
}

# ----------------------------------------------------------------------------
# EL JSON
# ----------------------------------------------------------------------------
# Mateix format estandard que la resta d'ESTRUCTURALS: familia 'capcalera',
# una seccio per bloc i un fill per linia.
#   tipus 'etiqueta' -> titol = l'etiqueta, cos = el valor
#   tipus 'text'     -> cos = el paragraf sencer
#   tipus 'buida'    -> una linia en blanc (hi es per no perdre la posicio)
function _CapModelAJson($blocs) {
    $nodes = New-Object System.Collections.ArrayList
    foreach ($b in @($blocs)) {
        $fills = New-Object System.Collections.ArrayList
        foreach ($l in @($b.Linies)) {
            [void]$fills.Add([ordered]@{
                tipus = [string]$l.Tipus
                titol = [string]$l.Etiqueta
                clau  = ('p' + [string]$l.Para)
                cos   = @([ordered]@{ runs = @([ordered]@{ t = [string]$l.Valor }) })
                fills = @()
            })
        }
        [void]$nodes.Add([ordered]@{
            tipus = 'seccio'
            titol = (_CapTitolDe ([string]$b.Clau))
            clau  = [string]$b.Clau
            cos   = @([ordered]@{ runs = @([ordered]@{ t = ("S'aplica a: " + ((_CapAplicaDe ([string]$b.Clau)) -join ', ')) }) })
            fills = $fills.ToArray()
        })
    }
    return [ordered]@{
        tipus   = 'capcalera'
        familia = 'capcalera'
        intro   = @()
        nodes   = $nodes.ToArray()
    }
}

# El JSON de la capcalera, generat del .docx.
function Build-CapcaleraJson([string]$docxPath = '') {
    if ([string]::IsNullOrWhiteSpace($docxPath)) { $docxPath = Get-CapcaleraDocxPath }
    $xml = _CapLlegeixDocumentXml $docxPath
    if ($null -eq $xml) { return $null }
    return (_CapModelAJson (_CapBlocsDelXml $xml))
}

# ----------------------------------------------------------------------------
# TORNAR-HO A ESCRIURE AL .docx (edicions de TEXT, mai un serialitzador)
# ----------------------------------------------------------------------------
# UN PARAGRAF DE LINIA amb l'etiqueta i el valor que se li diguin. Funcio PURA:
# rep el XML d'un paragraf de MOSTRA i en retorna un de nou.
#
# AIXI ES COM ES DERIVA EL .docx DEL JSON: no s'inventa cap paragraf -construir
# un <w:p> a ma voldria dir reproduir la sagnia, la parada de tabulacio, la
# lletra i el rsid de la plantilla- sino que es CLONA una linia que ja hi es i
# nomes se li canvia el text. Aixi el format no es pot perdre.
function _CapParagrafAmb([string]$pXml, [string]$etiqueta, [string]$valor, [string]$lletra = '') {
    $runs = @(_CapRunsDeParagraf $pXml)
    if ($runs.Count -eq 0) { return $pXml }
    # Els dos grups: la NEGRETA del principi (l'etiqueta) i la resta (el valor).
    # El tabulador del mig, si n'hi ha, es queda com esta.
    $iTab = -1
    for ($i = 0; $i -lt $runs.Count; $i++) { if ($runs[$i].EsTab) { $iTab = $i; break } }
    $grupEti = New-Object System.Collections.ArrayList
    $grupVal = New-Object System.Collections.ArrayList
    if ($iTab -ge 0) {
        for ($i = 0; $i -lt $iTab; $i++) { [void]$grupEti.Add($runs[$i]) }
        for ($i = $iTab + 1; $i -lt $runs.Count; $i++) { [void]$grupVal.Add($runs[$i]) }
    } else {
        $i = 0
        while ($i -lt $runs.Count -and [bool]$runs[$i].Negreta) { [void]$grupEti.Add($runs[$i]); $i++ }
        for (; $i -lt $runs.Count; $i++) { [void]$grupVal.Add($runs[$i]) }
    }
    $canvis = New-Object System.Collections.ArrayList
    foreach ($c in @(_CapCanvisDeGrupText $grupEti.ToArray() 0 $etiqueta)) { [void]$canvis.Add($c) }
    foreach ($c in @(_CapCanvisDeGrupText $grupVal.ToArray() 0 $valor))    { [void]$canvis.Add($c) }
    $out = [string]$pXml
    foreach ($c in @(@($canvis) | Sort-Object -Property @{ Expression = { [int]$_.Ini }; Descending = $true })) {
        $out = $out.Substring(0, [int]$c.Ini) + [string]$c.Nou + $out.Substring([int]$c.Fi)
    }
    if (-not [string]::IsNullOrWhiteSpace($lletra)) { $out = _CapAplicaLletra $out $lletra }
    return $out
}

# LA LLETRA DE LA CAPCALERA, de la configuracio. Funcio PURA.
#
# PER QUE: el .docx porta el tipus de lletra escrit a CADA <w:r>, o sigui que
# canviar $ReportFormatConfig.BodyFontName no en canviava res -la capcalera era
# l'unic tros de l'informe que no obeia la configuracio-. Com que ara les linies
# es generen, aqui es on es tanca.
function _CapAplicaLletra([string]$pXml, [string]$lletra) {
    if ([string]::IsNullOrWhiteSpace($lletra)) { return $pXml }
    $l = (_CapEscapa $lletra)
    return [regex]::Replace($pXml, '<w:rFonts\b[^>]*/>', ('<w:rFonts w:ascii="' + $l + '" w:hAnsi="' + $l + '"/>'))
}

# LA REGIO DE LINIES d'un bloc: del primer paragraf d'etiqueta a l'ultim. Es el
# tros que es refa des del JSON; el que hi ha a fora (el departament, la linia
# en blanc i l'INFORME) no es toca. Funcio PURA: retorna @{ Ini; Fi } amb
# INDEXS de paragraf, o $null si el bloc no en te cap.
function _CapRegioDeBloc($linies) {
    $ini = -1; $fi = -1
    foreach ($l in @($linies)) {
        if ([string]$l.Tipus -ne 'etiqueta') { continue }
        if ($ini -lt 0) { $ini = [int]$l.Para }
        $fi = [int]$l.Para
    }
    if ($ini -lt 0) { return $null }
    return @{ Ini = $ini; Fi = $fi }
}

# Aplica el JSON al XML del document. Funcio PURA: retorna el XML nou.
#
# EL JSON MANA I EL .docx ES DERIVA. De cada bloc se'n refa la REGIO DE LINIES
# (del primer camp a l'ultim) amb el que digui el JSON: quines linies hi ha, en
# quin ordre i amb quina etiqueta. Aixi afegir la Classificacio a la capcalera
# generica -o treure les Dates de la d'activitats extraordinaries- es editar el
# JSON, i no hi ha tres blocs gairebe iguals per mantenir a ma.
#
# La resta del document (el departament, la linia en blanc, l'INFORME i els
# marcadors [[CAP:X]]) no es toca mai.
function _CapAplicaAlXml([string]$xml, $json, [string]$lletra = '') {
    $paras = @(_CapTrossejaParagrafs $xml)
    $blocs = @(_CapBlocsDelXml $xml)
    # El JSON, per clau de bloc.
    $perClau = @{}
    foreach ($n in @($json.nodes)) { $perClau[[string]$n.clau] = $n }

    # De DARRERE CAP ENDAVANT: refer una regio mou els offsets de tot el que ve
    # despres.
    $regions = New-Object System.Collections.ArrayList
    foreach ($b in $blocs) {
        $nd = $null
        if ($perClau.ContainsKey([string]$b.Clau)) { $nd = $perClau[[string]$b.Clau] }
        if ($null -eq $nd) { continue }
        $reg = _CapRegioDeBloc $b.Linies
        if ($null -eq $reg) { continue }
        # La MOSTRA: la primera linia d'etiqueta del bloc.
        $mostra = [string]$paras[[int]$reg.Ini].Xml
        # Les linies que vol el JSON (nomes camps i blancs de dins de la regio).
        $vol = New-Object System.Collections.ArrayList
        foreach ($f in @($nd.fills)) {
            $tp = [string]$f.tipus
            if ($tp -ne 'etiqueta' -and $tp -ne 'buida') { continue }
            $valor = ''
            foreach ($pp in @($f.cos)) { foreach ($r in @($pp.runs)) { $valor += [string]$r.t } }
            [void]$vol.Add(@{ Tipus = $tp; Etiqueta = [string]$f.titol; Valor = $valor })
        }
        # LA REGIO ES DEL PRIMER CAMP A L'ULTIM, tambe al JSON: les linies en
        # blanc de fora (la que separa els camps de l'INFORME) no en formen
        # part. Sense retallar-les, la llista del JSON i la del document no
        # quadraven mai i la regio es refeia sempre.
        $iniV = -1; $fiV = -1
        for ($i = 0; $i -lt $vol.Count; $i++) {
            if ([string]$vol[$i].Tipus -ne 'etiqueta') { continue }
            if ($iniV -lt 0) { $iniV = $i }
            $fiV = $i
        }
        if ($iniV -lt 0) { continue }
        $vol = [System.Collections.ArrayList]@($vol[$iniV..$fiV])
        if ($vol.Count -eq 0) { continue }

        # QUE HI HA ARA a la regio, per no refer-la si no cal.
        $ara = New-Object System.Collections.ArrayList
        for ($i = [int]$reg.Ini; $i -le [int]$reg.Fi; $i++) {
            $lin = _CapLiniaDeRuns (_CapRunsDeParagraf ([string]$paras[$i].Xml))
            [void]$ara.Add(@{ Tipus = [string]$lin.Tipus; Etiqueta = [string]$lin.Etiqueta; Valor = [string]$lin.Valor; Para = $i })
        }
        # MATEIXA ESTRUCTURA? Llavors nomes cal canviar EL TEXT alli on estigui,
        # i el document no es toca mes. Refer-la sencera funcionaria igual, pero
        # deixaria el .docx diferent a cada desat encara que no s'hagues canviat
        # res -i llavors 'Actualitzar.bat' committejaria una capcalera "nova"
        # cada vegada-.
        $mateixa = ($ara.Count -eq $vol.Count)
        if ($mateixa) {
            for ($i = 0; $i -lt $vol.Count; $i++) {
                if ([string]$ara[$i].Tipus -ne [string]$vol[$i].Tipus) { $mateixa = $false; break }
            }
        }
        if ($mateixa) {
            for ($i = 0; $i -lt $vol.Count; $i++) {
                if ([string]$vol[$i].Tipus -eq 'buida') { continue }
                $iP = [int]$ara[$i].Para
                if ([string]$ara[$i].Etiqueta -eq [string]$vol[$i].Etiqueta -and
                    [string]$ara[$i].Valor    -eq [string]$vol[$i].Valor) { continue }
                [void]$regions.Add(@{
                    Ini = [int]$paras[$iP].Ini
                    Fi  = [int]$paras[$iP].Fi
                    Nou = (_CapParagrafAmb ([string]$paras[$iP].Xml) ([string]$vol[$i].Etiqueta) ([string]$vol[$i].Valor) '')
                })
            }
            continue
        }

        # L'ESTRUCTURA HA CANVIAT (s'hi ha afegit o tret una linia, o s'han
        # reordenat): es refa la regio sencera clonant la linia de mostra.
        $nouXml = ''
        foreach ($v in $vol) {
            if ([string]$v.Tipus -eq 'buida') {
                $nouXml += (_CapParagrafAmb $mostra '' '' $lletra)
            } else {
                $nouXml += (_CapParagrafAmb $mostra ([string]$v.Etiqueta) ([string]$v.Valor) $lletra)
            }
        }
        [void]$regions.Add(@{
            Ini = [int]$paras[[int]$reg.Ini].Ini
            Fi  = [int]$paras[[int]$reg.Fi].Fi
            Nou = $nouXml
        })
    }
    $out = [string]$xml
    foreach ($r in @(@($regions) | Sort-Object -Property @{ Expression = { [int]$_.Ini }; Descending = $true })) {
        $out = $out.Substring(0, [int]$r.Ini) + [string]$r.Nou + $out.Substring([int]$r.Fi)
    }
    return $out
}

# Reescriu els <w:t> d'un grup de runs: TOT el text al primer, els altres buits.
# El <w:rPr> (lletra, negreta, mida) no es toca mai. Funcio PURA.
function _CapCanvisDeGrupText($runs, [int]$offsetPara, [string]$text) {
    $canvis = New-Object System.Collections.ArrayList
    $rs = @($runs)
    $posat = $false
    foreach ($r in $rs) {
        $runXml = [string]$r.Xml
        if (-not [regex]::IsMatch($runXml, '<w:t(?:\s[^>]*)?>')) { continue }
        $nouText = if (-not $posat) { (_CapEscapa $text) } else { '' }
        $posat = $true
        $primer = $true
        $nouRun = [regex]::Replace($runXml, '<w:t(?:\s[^>]*)?>(.*?)</w:t>', {
            param($m)
            if ($primer) { $primer = $false; return ('<w:t xml:space="preserve">' + $nouText + '</w:t>') }
            return '<w:t xml:space="preserve"></w:t>'
        }, 'Singleline')
        if ($nouRun -eq $runXml) { continue }
        [void]$canvis.Add(@{ Ini = ($offsetPara + [int]$r.Ini); Fi = ($offsetPara + [int]$r.Fi); Nou = $nouRun })
    }
    return $canvis.ToArray()
}

# ----------------------------------------------------------------------------
# ZIP: llegir i escriure 'word/document.xml' sense tocar res mes
# ----------------------------------------------------------------------------
function _CapLlegeixDocumentXml([string]$docxPath) {
    if (-not (Test-Path -LiteralPath $docxPath)) { return $null }
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
        $e = $zip.GetEntry('word/document.xml')
        if ($null -eq $e) { return $null }
        $sr = New-Object System.IO.StreamReader($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
        try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
    } catch {
        return $null
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

# Escriu el XML nou DINS del .docx. Nomes es reescriu aquesta entrada; l'ordre i
# la resta del ZIP no es toquen (mode 'Update').
function _CapEscriuDocumentXml([string]$docxPath, [string]$xml) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::Open($docxPath, [System.IO.Compression.ZipArchiveMode]::Update)
        $e = $zip.GetEntry('word/document.xml')
        if ($null -eq $e) { return $false }
        $st = $e.Open()
        try {
            $st.SetLength(0)
            $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($xml)
            $st.Write($bytes, 0, $bytes.Length)
        } finally { $st.Dispose() }
        return $true
    } catch {
        return $false
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

# ----------------------------------------------------------------------------
# LES DUES OPERACIONS QUE FA SERVIR EL PROGRAMA
# ----------------------------------------------------------------------------
# Genera (o refresca) '0 CAPCALERA.json' a partir del .docx. Es crida en obrir
# l'editor: aixi el JSON mai pot quedar desincronitzat del document.
function Sync-CapcaleraJson([string]$docxPath = '', [string]$jsonPath = '') {
    if ([string]::IsNullOrWhiteSpace($docxPath)) { $docxPath = Get-CapcaleraDocxPath }
    if ([string]::IsNullOrWhiteSpace($jsonPath)) { $jsonPath = Get-CapcaleraJsonPath }
    $model = Build-CapcaleraJson $docxPath
    if ($null -eq $model) { return $false }
    ($model | ConvertTo-Json -Depth 20) | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    return $true
}

# ...i al reves: el que digui el JSON, cap al .docx. Retorna @{ Ok; Canvis }.
function Apply-CapcaleraJson([string]$jsonPath = '', [string]$docxPath = '') {
    if ([string]::IsNullOrWhiteSpace($docxPath)) { $docxPath = Get-CapcaleraDocxPath }
    if ([string]::IsNullOrWhiteSpace($jsonPath)) { $jsonPath = Get-CapcaleraJsonPath }
    if (-not (Test-Path -LiteralPath $jsonPath)) { return @{ Ok = $false; Motiu = 'no hi ha el JSON' } }
    $xml = _CapLlegeixDocumentXml $docxPath
    if ($null -eq $xml) { return @{ Ok = $false; Motiu = 'no s''ha pogut llegir el document' } }
    $json = $null
    try { $json = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { }
    if ($null -eq $json) { return @{ Ok = $false; Motiu = 'el JSON no es valid' } }
    # LA LLETRA, de la configuracio de format: la capcalera era l'unic tros de
    # l'informe que no l'obeia (la porta escrita a cada <w:r> del .docx).
    $lletra = ''
    try { $lletra = [string]$Script:ReportFormatConfig.BodyFontName } catch { }
    $nou = _CapAplicaAlXml $xml $json $lletra
    if ($nou -eq $xml) { return @{ Ok = $true; Canvis = $false } }
    # COPIA DE SEGURETAT abans de tocar l'unica plantilla que no es pot refer.
    try { Copy-Item -LiteralPath $docxPath -Destination ($docxPath + '.bak') -Force } catch { }
    if (-not (_CapEscriuDocumentXml $docxPath $nou)) { return @{ Ok = $false; Motiu = 'no s''ha pogut desar el document' } }
    return @{ Ok = $true; Canvis = $true }
}
