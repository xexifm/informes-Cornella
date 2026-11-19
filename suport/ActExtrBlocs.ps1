#requires -Version 5.1
<#
.SYNOPSIS
  Mode ACT_EXTR: el requeriment i l'informe favorable en blocs
  (Build-ActExtrBlocs, PURA) i el .docx (Build-ActExtrDocument).

  Part del modul ACT_EXTR, partit per responsabilitats (com Llicencia):
    ActExtrDades.ps1     logica del Decret, punts, plantilla i registre (PURES)
    ActExtrBlocs.ps1     composicio del document en blocs (PURA) + el .docx
    ActExtrPantalles.ps1 les finestres (WinForms)
    ActExtr.ps1          l'orquestrador (Invoke-ActExtrFlow) i la descripcio
  Tot va amb dot-source al mateix ambit (Motor.ps1 els carrega en aquest ordre).
#>

# ----------------------------------------------------------------------------
# Composicio del document (requeriment o informe favorable)
# ----------------------------------------------------------------------------
# Capcalera ACT_EXTR: hashtable amb les claus que espera Apply-HeaderReplacements
# mes <<DATES>> i <<AFORAMENT>>.
function _ActExtrHeaderMap($header) {
    $get = {
        param($k)
        if ($null -eq $header) { return '' }
        if ($header -is [System.Collections.IDictionary]) { if ($header.Contains($k)) { return [string]$header[$k] } ; return '' }
        if ($header.PSObject.Properties.Name -contains $k) { return [string]$header.$k }
        return ''
    }
    return @{
        ID_GIA    = (& $get 'ID_GIA')
        EXP_NUM   = (& $get 'EXP_NUM')
        ADRECA    = (& $get 'ADRECA')
        ACTIVITAT = (& $get 'ACTIVITAT')
        TITULAR   = (& $get 'TITULAR')
        DATES     = (& $get 'DATES')
        AFORAMENT = (& $get 'AFORAMENT')
        # Claus de REQ1 que la capcalera ACT_EXTR no usa (per si de cas, buides).
        NUM_ANOTACIO  = ''
        DATA_ANOTACIO = ''
    }
}

# Nom del fitxer de sortida: YYYY-MM-DD_ActExtr-<Tipus>_GIA <id>_<Activitat>.docx
# S'hi inclou el nom de l'ACTIVITAT perque en un mateix establiment (mateix GIA)
# es poden fer diferents activitats extraordinaries i s'han de poder distingir.
function _GetActExtrOutputFileName([string]$tipus, [string]$gia, [string]$activitat) {
    $today = (Get-Date).ToString('yyyy-MM-dd')
    if ([string]::IsNullOrWhiteSpace($gia)) { $gia = 's_n' }
    $gia = ($gia -replace '[\\/:*?"<>|]','_').Trim()
    $name = "{0}_ActExtr-{1}_GIA {2}" -f $today, $tipus, $gia
    $act = ([string]$activitat -replace '[\\/:*?"<>|]','_').Trim()
    if (-not [string]::IsNullOrWhiteSpace($act)) { $name += "_$act" }
    return ($name + '.docx')
}

# Emet el cos del document a partir dels blocs de la plantilla. Segons el
# marcador Titol 2 (Kind), cada paragraf de contingut es renderitza diferent:
#   'item'   -> item: NUMERAT al requeriment; amb PIC (vinyeta) al favorable.
#   'child'  -> sub-apartat amb PIC i sagnat.
#   'note'   -> sub-paragraf sagnat SENSE pic (nota).
#   'text'   -> paragraf de cos (sense pic ni numero).
#   'header' -> capcalera de conclusions (centrada i en negreta).
#   'conc'   -> paragraf de conclusio (justificat).
#
# ESPAIAT (diferent en cada mode, per reproduir el format de referencia):
#   - REQUERIMENT: cada unitat (item o paragraf) va separada per una linia en
#     blanc (com REQ1). Sub-items i URLs s'enganxen a la unitat anterior.
#   - FAVORABLE: els punts d'una mateixa seccio se separen amb SpaceBefore (no
#     linies en blanc); nomes s'insereix una linia en blanc al CANVI de seccio
#     (Titol 1 de la plantilla). Aixi la llista surt compacta, com el document
#     de referencia.
function _WriteActExtrBody($sel, $blocks, $mode, $ctx, $computed) {
    [void](Write-Informe $sel @(Build-ActExtrBlocs $blocks $mode $ctx $computed))
}


# EL COS D'ACT_EXTR EN BLOCS. Funcio PURA (es prova a Linux comptant blocs); qui
# ho escriu es Write-Informe, com a la resta d'informes. Abans aquestes dues
# regles escrivien directament al Word, i eren les ultimes del programa.
#
# El PRIMER SUB-PUNT d'una unitat (12 pt en lloc de 6) aqui el decideix aquesta
# funcio i el passa al bloc ('First'): a ACT_EXTR obre unitat QUALSEVOL text que
# no sigui un sub-punt, i el motor nomes ho sap fer amb blocs 'unitat'.
function Build-ActExtrBlocs($blocks, $mode, $ctx, $computed) {
    if ($mode -eq 'fav') { return @(_ActExtrBlocsFav $blocks $ctx $computed) }
    if ($mode -eq 'no')  { return @(Build-ActExtrNoBlocs $blocks) }

    # ---- REQUERIMENT ----
    $out = New-Object System.Collections.ArrayList
    $num0 = 0   # comptador d'items de 1r nivell (continu)
    $first = $true
    $primerFill = $false   # el seguent sub-punt es el primer de la seva unitat?
    foreach ($block in $blocks) {
        if (-not (Test-ActExtrIncludeBlock $block.Key $mode $ctx)) { continue }
        $kind = [string]$block.Kind
        foreach ($c in $block.Contents) {
            $resolved = Resolve-ActExtrTokens ([string]$c.Text) $computed

            if ($c.IsUrl) {
                $u = $resolved
                if ($u.StartsWith('[[URL]] ')) { $u = $u.Substring('[[URL]] '.Length).Trim() }
                if (-not [string]::IsNullOrWhiteSpace($u)) { [void]$out.Add(@{ T = 'enllac'; Url = $u; Fill = ($kind -eq 'child') }) }
                continue
            }

            $parts = _SplitTextAndUrls $resolved
            if ([string]::IsNullOrWhiteSpace($parts.Text) -and @($parts.Urls).Count -eq 0) { continue }

            if ($kind -eq 'child') {
                if (-not [string]::IsNullOrWhiteSpace($parts.Text)) {
                    [void]$out.Add(@{ T = 'pic'; Text = $parts.Text; Fill = $true; First = $primerFill })
                    $primerFill = $false
                }
                foreach ($x in $parts.Urls) { [void]$out.Add(@{ T = 'enllac'; Url = $x; Fill = $true }) }
                $first = $false
                continue
            }

            # Unitat nova (item o paragraf de cos): linia en blanc al DAVANT si
            # no es la primera unitat del document. La mateixa bandera que l'aire
            # entre items de REQ1 ('item'); aqui va davant i no darrere, i per
            # aixo el document no acaba amb un paragraf en blanc de mes.
            if (-not $first) { [void]$out.Add(@{ T = 'aire'; Clau = 'item' }) }
            if ($kind -eq 'item') {
                $num0++
                if (-not [string]::IsNullOrWhiteSpace($parts.Text)) { [void]$out.Add(@{ T = 'item'; Num = "$num0."; Text = $parts.Text }) }
            } else {
                # "Ho poso al seu coneixement..." sempre separat (Format.ps1).
                if (_EsFraseTancament $parts.Text) { [void]$out.Add(@{ T = 'separa' }) }
                if (-not [string]::IsNullOrWhiteSpace($parts.Text)) { [void]$out.Add(@{ T = 'cos'; Text = $parts.Text }) }
            }
            foreach ($x in $parts.Urls) { [void]$out.Add(@{ T = 'enllac'; Url = $x }) }
            # La unitat que ve de tancar-se es la mare dels sub-punts seguents.
            $primerFill = $true
            $first = $false
        }
    }
    return $out.ToArray()
}

# El cos de l'INFORME FAVORABLE en blocs (espaiat per seccions). Els punts van
# amb PIC i separats per SpaceBefore; entre seccions (canvi de Titol 1 de la
# plantilla) s'hi posa una linia en blanc. Funcio PURA.
function _ActExtrBlocsFav($blocks, $ctx, $computed) {
    $out = New-Object System.Collections.ArrayList
    $firstBlock  = $true
    $prevSection = -1
    $primerFill  = $false   # el seguent sub-punt es el primer de la seva unitat?
    foreach ($block in $blocks) {
        if (-not (Test-ActExtrIncludeBlock $block.Key 'fav' $ctx)) { continue }
        $kind = [string]$block.Kind
        # Linia en blanc nomes al canvi de seccio (no dins d'una seccio): la
        # mateixa bandera que separa les seccions a la resta d'informes.
        if ((-not $firstBlock) -and ([int]$block.Section -ne [int]$prevSection)) { [void]$out.Add(@{ T = 'aire'; Clau = 'seccio' }) }

        foreach ($c in $block.Contents) {
            $resolved = Resolve-ActExtrTokens ([string]$c.Text) $computed

            if ($c.IsUrl) {
                $u = $resolved
                if ($u.StartsWith('[[URL]] ')) { $u = $u.Substring('[[URL]] '.Length).Trim() }
                if (-not [string]::IsNullOrWhiteSpace($u)) { [void]$out.Add(@{ T = 'enllac'; Url = $u; Fill = ($kind -eq 'child') }) }
                continue
            }

            $parts = _SplitTextAndUrls $resolved
            if ([string]::IsNullOrWhiteSpace($parts.Text) -and @($parts.Urls).Count -eq 0) { continue }
            $txt = $parts.Text

            # "Ho poso al seu coneixement..." sempre separat (Format.ps1).
            if (_EsFraseTancament $txt) { [void]$out.Add(@{ T = 'separa' }) }
            # -First al PRIMER sub-punt de cada unitat (12 pt en lloc de 6).
            # Qualsevol contingut que NO sigui un sub-punt obre unitat nova.
            if ($kind -eq 'child') {
                if ($txt) {
                    [void]$out.Add(@{ T = 'pic'; Text = $txt; Fill = $true; First = $primerFill })
                    $primerFill = $false
                }
                foreach ($x in $parts.Urls) { [void]$out.Add(@{ T = 'enllac'; Url = $x; Fill = $true }) }
            } else {
                # Els enllacos d'una NOTA van sagnats (de fill); la resta, no.
                $urlFill = $false
                switch ($kind) {
                    'note'   { if ($txt) { [void]$out.Add(@{ T = 'nota'; Text = $txt }) }; $urlFill = $true }
                    'label'  { if ($txt) { [void]$out.Add(@{ T = 'etiqueta'; Text = $txt }) } }
                    'header' { if ($txt) { [void]$out.Add(@{ T = 'conclusiocap'; Text = $txt }) } }
                    'conc'   { if ($txt) { [void]$out.Add(@{ T = 'conclusio'; Text = $txt }) } }
                    'text'   { if ($txt) { [void]$out.Add(@{ T = 'cos'; Text = $txt }) } }
                    # 'item': pic de 1r nivell, que MAI porta la separacio gran.
                    default  { if ($txt) { [void]$out.Add(@{ T = 'pic'; Text = $txt; First = $false }) } }
                }
                # Un 'header' no escrivia els seus enllacos (sempre ha estat aixi).
                if ($kind -ne 'header') {
                    foreach ($x in $parts.Urls) { [void]$out.Add(@{ T = 'enllac'; Url = $x; Fill = $urlFill }) }
                }
                $primerFill = $true
            }
        }
        $prevSection = [int]$block.Section
        $firstBlock  = $false
    }
    return $out.ToArray()
}

# L'INFORME "NO ES UNA ACTIVITAT EXTRAORDINARIA" en blocs (l'usuari, novembre
# 2026, amb el model "Act_Extr_NOM": "nomes canvien les dades abans de la
# paraula INFORME, la resta es estatic"). La capcalera es la mateixa d'ACT_EXTR
# (0 CAPCALERA, bloc ACT_EXTR); el cos, TOT el de la plantilla ACT_EXTR_NO.json,
# sense cap regla del Decret. Funcio PURA.
#   - Una linia en blanc entre paragrafs, menys entre els pics d'una mateixa
#     llista (el model els porta seguits).
#   - Cada SECCIO de la plantilla a partir de la segona comenca en una PAGINA
#     NOVA: els motius van "a continuacio de la signatura electronica".
#   - 'item' = pic; 'capcalera' = titol centrat en negreta (INFORME); 'text' =
#     cos amb la seva **negreta** i //cursiva//; els enllacos, com a enllac.
function Build-ActExtrNoBlocs($blocks) {
    $out = New-Object System.Collections.ArrayList
    $prev = ''          # el tipus del darrer paragraf escrit ('' = cap a la pagina)
    $prevSection = $null
    $despresDeSalt = $false
    foreach ($block in @($blocks)) {
        if ($null -eq $block) { continue }
        $kind = [string]$block.Kind
        if ($null -ne $prevSection -and [int]$block.Section -ne [int]$prevSection -and $prev -ne '') {
            [void]$out.Add(@{ T = 'saltpagina' })
            $prev = ''; $despresDeSalt = $true
        }
        $prevSection = [int]$block.Section
        foreach ($c in @($block.Contents)) {
            $txt = [string]$c.Text
            if ($c.IsUrl) {
                $u = $txt
                if ($u.StartsWith('[[URL]] ')) { $u = $u.Substring('[[URL]] '.Length) }
                $u = $u.Trim()
                if ($u -eq '') { continue }
                if ($prev -ne '') { [void]$out.Add(@{ T = 'espai' }) }
                [void]$out.Add(@{ T = 'enllac'; Url = $u })
                $prev = 'url'; $despresDeSalt = $false
                continue
            }
            $parts = _SplitTextAndUrls $txt
            if ([string]::IsNullOrWhiteSpace($parts.Text) -and @($parts.Urls).Count -eq 0) { continue }
            # Els pics seguits, sense linia en blanc entre ells.
            if ($prev -ne '' -and -not ($kind -eq 'item' -and $prev -eq 'item')) { [void]$out.Add(@{ T = 'espai' }) }
            if (-not [string]::IsNullOrWhiteSpace($parts.Text)) {
                switch ($kind) {
                    'item'   { [void]$out.Add(@{ T = 'pic'; Text = $parts.Text; First = $false }) }
                    'child'  { [void]$out.Add(@{ T = 'pic'; Text = $parts.Text; Fill = $true; First = $false }) }
                    # El titol centrat ja va en negreta i s'escriu TAL QUAL
                    # (Format-ConclusionHeader, -Pla): sense les marques, que
                    # sortirien escrites ("**INFORME**").
                    'header' { [void]$out.Add(@{ T = 'conclusiocap'; Text = ($parts.Text -replace '\*\*|//', ''); DespresDeSalt = $despresDeSalt }) }
                    'note'   { [void]$out.Add(@{ T = 'nota'; Text = $parts.Text }) }
                    'label'  { [void]$out.Add(@{ T = 'etiqueta'; Text = $parts.Text }) }
                    'conc'   { [void]$out.Add(@{ T = 'conclusio'; Text = $parts.Text }) }
                    default  {
                        # "Ho poso al seu coneixement..." sempre separat (Format.ps1).
                        if (_EsFraseTancament $parts.Text) { [void]$out.Add(@{ T = 'separa' }) }
                        [void]$out.Add(@{ T = 'cos'; Text = $parts.Text; DespresDeSalt = $despresDeSalt })
                    }
                }
            }
            foreach ($x in @($parts.Urls)) { [void]$out.Add(@{ T = 'enllac'; Url = $x }) }
            $prev = $kind; $despresDeSalt = $false
        }
    }
    return $out.ToArray()
}

# Genera el document (requeriment, informe favorable o "no es una activitat
# extraordinaria") i retorna la ruta del .docx. $mode = 'req' | 'fav' | 'no'. Reutilitza la capcalera 0 CAPCALERA.docx
# (bloc ACT_EXTR) i el motor de format.
function Build-ActExtrDocument($word, $header, $decret, $delivered, $mode) {
    $computed = Get-ActExtrComputed $decret
    $status   = Get-ActExtrStatus $decret $computed $delivered
    $statusByKey = @{}
    foreach ($s in $status) { $statusByKey[$s.Key] = $s }
    $defKeys = Get-ActExtrDeficiencies $decret $computed $delivered
    $ctx = @{ Decret=$decret; Computed=$computed; Delivered=$delivered; StatusByKey=$statusByKey; DefKeys=$defKeys }

    $tplPath = switch ($mode) { 'fav' { $script:ActExtrFavTemplate } 'no' { $script:ActExtrNoTemplate } default { $script:ActExtrReqTemplate } }
    $blocks  = Parse-ActExtrTemplate $tplPath

    $tipus = switch ($mode) { 'fav' { 'Fav' } 'no' { 'NoExtr' } default { 'Req' } }
    $gia = if ($header -is [System.Collections.IDictionary]) { [string]$header['ID_GIA'] } else { [string]$header.ID_GIA }
    $act = if ($header -is [System.Collections.IDictionary]) { [string]$header['ACTIVITAT'] } else { [string]$header.ACTIVITAT }
    $baseName  = _GetActExtrOutputFileName $tipus $gia $act

    # Retallem la capcalera per quedar-nos nomes amb el bloc ACT_EXTR.
    # El cos hereta l'estil 'List Paragraph' del darrer paragraf d'aquella
    # capcalera (igual que REQ1), que ja resol a Bookman Old Style: el format
    # surt del motor Format.ps1 i de l'estil, no d'un override.
    return Write-InformeDocx $word $baseName 'ACT_EXTR' (_ActExtrHeaderMap $header) {
        param($sel)
        _WriteActExtrBody $sel $blocks $mode $ctx $computed
    }
}
