#requires -Version 5.1
<#
.SYNOPSIS
  "Actualitzar base": recorre la carpeta d'informes i en fa la base
  (local\base-dades-activitats\informes-db.json).

.DESCRIPTION
  Vivia a Informes.ps1, que tambe te l'editor de la base i els classificadors
  del text: amb les edicions a ma i el mode automatic (octubre 2026) passava de
  les 1.200 linies. Aqui hi ha el que LLEGEIX els informes del disc i munta la
  base; a Informes.ps1, el que n'enten el text (conclusio, GIA, expedient) i
  l'editor.

  NOMES DEFINEIX FUNCIONS.
#>

# ----------------------------------------------------------------------------
# Lectura del .docx (necessita les primitives de Seguiment.ps1; no headless)
# ----------------------------------------------------------------------------
# Retorna un array amb el text de TOTS els paragrafs del document, INCLOENT els
# de dins de taules (la capcalera amb ID GIA / Exp. Num viu en una taula, i
# _BodyParagraphsXml les salta; per aixo seleccionem './/w:p').
function _ReadDocxParagraphs($docxPath) {
    $info = _LoadDocxXml $docxPath
    $out = New-Object System.Collections.ArrayList
    foreach ($p in $info.Body.SelectNodes('.//w:p', $info.Ns)) {
        [void]$out.Add((_ParagraphTextXml $p $info.Ns))
    }
    return $out.ToArray()
}

# Llegeix tots els paragrafs d'un .doc antic (Word 97-2003) via Word COM, en
# nomes-lectura. Necessita una instancia de Word JA OBERTA ($wordApp, creada
# mandrosament nomes si cal a Invoke-InformesDbScan); aqui nomes s'obre i es
# tanca el DOCUMENT (mai l'aplicacio).
function _ReadDocParagraphsWord($wordApp, $docPath) {
    $doc = $wordApp.Documents.Open($docPath, $false, $true, $false)
    try {
        $out = New-Object System.Collections.ArrayList
        foreach ($p in $doc.Paragraphs) {
            $t = $p.Range.Text
            if ($null -ne $t) { $t = $t.TrimEnd([char]13, [char]7) }
            [void]$out.Add($t)
        }
        return $out.ToArray()
    } finally {
        $doc.Close($false)
    }
}

# Tria com llegir els paragrafs d'un informe segons l'extensio: .docx (sense
# Word, via zip) o .doc antic (via Word COM; retorna buit si no hi ha Word
# disponible, i l'informe queda "a revisar" com si no s'hagues pogut llegir).
function _ReadInformeParagraphs($file, $wordApp) {
    if ($file.Extension -ieq '.doc') {
        if ($null -eq $wordApp) { return @() }
        return _ReadDocParagraphsWord $wordApp $file.FullName
    }
    return _ReadDocxParagraphs $file.FullName
}

# ----------------------------------------------------------------------------
# ON es busquen els informes: NOMES el primer nivell
# ----------------------------------------------------------------------------
# Els fitxers que son DIRECTAMENT dins de cada carpeta d'activitat de la
# carpeta d'informes (<arrel>\<carpeta>\fitxer). Ni els de l'arrel ni els de
# les subcarpetes (l'usuari, octubre 2026: "nomes ha de tenir en compte els
# informes de les carpetes dins la carpeta d'informes, pero no dins les
# subcarpetes"). A les subcarpetes hi ha l'expedient del GIA descarregat
# (~25.000 fitxers) i informes vells o d'altres activitats que no han de
# decidir l'estat; abans un Get-ChildItem -Recurse els recorria tots. El fan
# servir la base d'informes i el repas de contactes.
function Get-FitxersPrimerNivell([string]$dir) {
    $out = New-Object System.Collections.ArrayList
    foreach ($d in @(Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue)) {
        foreach ($f in @(Get-ChildItem -LiteralPath $d.FullName -File -ErrorAction SilentlyContinue)) { [void]$out.Add($f) }
    }
    return ,$out.ToArray()
}

# La clau (_ClauInforme) es d'un fitxer de primer nivell? "carpeta\fitxer",
# exactament una barra. PURA.
function _EsDePrimerNivell([string]$clau) {
    return (([string]$clau -replace '/', '\').Trim('\').Split('\').Count -eq 2)
}

# ----------------------------------------------------------------------------
# L'ID GIA ASSIGNAT A MA (local\base-dades-activitats\gia-assignats_*.json)
# ----------------------------------------------------------------------------
# Els informes antics no porten l'ID GIA a la capcalera, o el porten malament
# (els dos d'un titular deien 239 i eren del 293). Es van llegir un per un i
# se'n va desar l'assignacio a un fitxer de local\ (porta dades personals: com
# la classificacio, no es puja mai). Format:
#   { "informes":  [ { "ruta_relativa", "id_gia", "confianca", "motiu", "forcar" } ],
#     "sense_gia": [ { "ruta_relativa", "motiu" } ] }
# ruta_relativa: relativa a la carpeta d'informes, la mateixa clau que les
# correccions a ma (_ClauInforme).
#
# L'ORDRE de l'ID GIA d'un informe (Get-InformeData):
#   forcar=true > document > assignat > carpeta "GIA n" > Excel per expedient
# (i, despres, el dels germans de carpeta). 'forcar' es per a una capcalera amb
# l'ID EQUIVOCAT; sense, el document mana i l'assignacio nomes omple el buit.

# El fitxer d'assignacions MES RECENT (pel nom, que porta la data), o $null.
function Find-GiaAssignats {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return $null }
    $c = @(Get-ChildItem -LiteralPath $LocalActivitatsDir -Filter 'gia-assignats_*.json' -File -ErrorAction SilentlyContinue |
           Sort-Object Name -Descending)
    if ($c.Count -eq 0) { return $null }
    return $c[0]
}

# El contingut del fitxer (ja llegit) -> @{ Mapa = clau -> @{ Gia; Forcar;
# Confianca; Motiu; Ruta }; SenseGia = clau -> motiu }. Les entrades sense ruta
# o sense id_gia no compten. PURA.
function _GiaAssignatsMapa($obj) {
    $mapa = @{}; $sense = @{}
    foreach ($e in @(_PropInf $obj 'informes')) {
        if ($null -eq $e) { continue }
        $ruta = [string](_PropInf $e 'ruta_relativa')
        $gia = ([string](_PropInf $e 'id_gia')).Trim()
        if ([string]::IsNullOrWhiteSpace($ruta) -or $gia -eq '') { continue }
        $mapa[(_ClauInforme $ruta '').TrimStart('\')] = @{
            Gia = $gia; Forcar = [bool](_PropInf $e 'forcar'); Ruta = $ruta
            Confianca = [string](_PropInf $e 'confianca'); Motiu = [string](_PropInf $e 'motiu')
        }
    }
    foreach ($e in @(_PropInf $obj 'sense_gia')) {
        if ($null -eq $e) { continue }
        $ruta = [string](_PropInf $e 'ruta_relativa')
        if ([string]::IsNullOrWhiteSpace($ruta)) { continue }
        $sense[(_ClauInforme $ruta '').TrimStart('\')] = [string](_PropInf $e 'motiu')
    }
    return @{ Mapa = $mapa; SenseGia = $sense }
}

# Llegeix el fitxer d'assignacions ($path; buit = el mes recent, Find-GiaAssignats).
# Torna @{ Mapa; SenseGia; Fitxer; Signatura; Error }. MAI NO PETA: un fitxer
# que no es pot llegir dona les llistes buides i l'Error (l'escaneig continua i
# ho diu). La Signatura (nom|data|mida) es desa a la base: si canvia, l'escaneig
# torna a llegir tots els informes (si no, una assignacio nova nomes arribaria
# als informes que algu tornes a desar).
function Read-GiaAssignats([string]$path = '') {
    $buit = @{ Mapa = @{}; SenseGia = @{}; Fitxer = ''; Signatura = ''; Error = '' }
    $f = $null
    if ([string]::IsNullOrWhiteSpace($path)) { $f = Find-GiaAssignats }
    elseif (Test-Path -LiteralPath $path) { $f = Get-Item -LiteralPath $path }
    else { $buit.Error = "No trobo el fitxer d'assignacions de GIA: $path"; return $buit }
    if ($null -eq $f) { return $buit }
    $buit.Fitxer = $f.FullName
    $buit.Signatura = $f.Name + '|' + $f.LastWriteTimeUtc.Ticks + '|' + $f.Length
    try {
        # Read-JsonFile torna $null (no llanca) si el fitxer no es un JSON valid:
        # sense aquesta comprovacio, un fitxer trencat deixaria d'assignar res
        # en silenci.
        $obj = Read-JsonFile $f.FullName
        if ($null -eq $obj -or ($null -eq $obj.PSObject.Properties['informes'] -and $null -eq $obj.PSObject.Properties['sense_gia'])) {
            $buit.Error = $f.Name + " es buit, no es un JSON valid o no porta 'informes': no s'hi ha assignat cap GIA."
            return $buit
        }
        $m = _GiaAssignatsMapa $obj
        $buit.Mapa = $m.Mapa; $buit.SenseGia = $m.SenseGia
    } catch {
        $buit.Error = "No s'ha pogut llegir " + $f.Name + ': ' + $_.Exception.Message
    }
    return $buit
}

# Les claus del fitxer d'assignacions (informes i sense_gia) que ja no troben
# cap informe a la carpeta ($claus: hashtable de les claus escanejades). PURA.
function _GiaAssignatsNoTrobats($assignats, $claus) {
    $out = New-Object System.Collections.ArrayList
    foreach ($k in @($assignats.Mapa.Keys) + @($assignats.SenseGia.Keys)) {
        if (-not $claus.ContainsKey($k)) { [void]$out.Add([string]$k) }
    }
    return @($out | Sort-Object -Unique)
}

# Analitza UN informe. Retorna un PSCustomObject amb data, gia, expedient,
# conclusio, tipus, fitxer, ruta, carpeta i els motius (si cal revisar-lo).
# $wordApp es opcional (nomes cal per als .doc antics; vegeu
# _ReadInformeParagraphs). El que se'n treu del text ho decideix
# _ClassificaInforme (InformesClassificacio.ps1), el mateix que fa servir
# ValidarClassificacio.ps1. $assignat: l'entrada del fitxer d'assignacions
# d'aquest informe (@{ Gia; Forcar }) o $null; vegeu l'ORDRE a dalt.
function Get-InformeData($file, $expToGia, $cache, $wordApp = $null, $assignat = $null) {
    $data = _ParseDataInformeFromName $file.Name
    $lines = @()
    try { $lines = @(_ReadInformeParagraphs $file $wordApp) } catch { $lines = @() }

    $gia = _ExtractIdGia $lines
    $giaCarpeta = _GiaFromFolderName $file.FullName
    $exp = _ExtractExpedient $lines
    $font = 'document'
    $giaAssignat = if ($null -ne $assignat) { [string]$assignat.Gia } else { '' }
    if ($giaAssignat -ne '' -and ([bool]$assignat.Forcar -or [string]::IsNullOrWhiteSpace($gia))) {
        $gia = $giaAssignat; $font = 'assignat'
    }
    if ([string]::IsNullOrWhiteSpace($gia)) {
        $gia = $giaCarpeta
        if (-not [string]::IsNullOrWhiteSpace($gia)) { $font = 'carpeta' }
    }
    if ([string]::IsNullOrWhiteSpace($gia) -and $null -ne $expToGia) {
        $key = _NormalitzaExpedient $exp
        if ($key -ne '' -and $expToGia.ContainsKey($key)) { $gia = $expToGia[$key]; $font = 'excel' }
    }

    $cl = _ClassificaInforme $lines $file.Name $exp

    $motius = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($gia)) { [void]$motius.Add('sense ID GIA') }
    # "GIA 101" a la carpeta i "ID GIA: 110" a la capcalera (les xifres girades): l'informe se
    # n'anava en silenci a l'activitat d'un altre titular. Es queda amb el del
    # document (es el que s'ha escrit per a aquell informe), pero es diu. Un GIA
    # assignat (forcat o no) ja s'ha mirat a ma: no se'n diu res.
    if ($font -eq 'document' -and $giaCarpeta -ne '' -and $giaCarpeta -ne [string]$gia) {
        [void]$motius.Add('GIA del document diferent del de la carpeta')
    }
    foreach ($m in @($cl.Motius)) { [void]$motius.Add([string]$m) }

    $titular = ''
    if ($null -ne $cache -and -not [string]::IsNullOrWhiteSpace($gia) -and $cache.ById.ContainsKey([string]$gia)) {
        $titular = [string]$cache.ById[[string]$gia].TITULAR
    }

    return [pscustomobject]@{
        Data          = $data
        Gia           = $gia
        GiaFont       = $font
        Expedient     = $exp
        Titular       = $titular
        Conclusio     = $cl.Conclusio
        ConclusioBreu = $cl.Breu
        Tipus         = $cl.Tipus
        Fitxer        = $file.Name
        Ruta          = $file.FullName
        Carpeta       = _CarpetaActivitat $file.FullName
        Modificat     = $file.LastWriteTimeUtc.ToString('o')
        Ignorat       = $false
        Motius        = $motius.ToArray()
        EditatAMa         = $false
        AutoConclusioBreu = ''
        AutoIgnorat       = $false
    }
}

# Decideix (funcio PURA) si un informe s'ha de tornar a parsejar (obrir el .docx)
# o si es pot reutilitzar l'entrada de l'escaneig anterior:
#   - Si NO en teniem entrada -> cal parsejar (es nou).
#   - Si en teniem i el fitxer s'ha modificat DESPRES de l'ultima actualitzacio
#     -> cal parsejar.
#   - Si en teniem i no s'ha tocat des de l'ultima actualitzacio -> reutilitzar.
# $lwUtc i $prevUtc son [datetime] en UTC.
function _HaDeReprocessar([datetime]$lwUtc, [datetime]$prevUtc, [bool]$teEntrada) {
    if (-not $teEntrada) { return $true }
    return ($lwUtc -gt $prevUtc)
}

# Aplana la base d'informes carregada (objecte de ConvertFrom-Json) en registres
# plans indexats per la ruta RELATIVA a la seva carpeta_arrel (_ClauInforme),
# arrossegant les dades de l'activitat a cada informe. Cada registre te la
# mateixa forma que Get-InformeData (perque el reagrupament els tracti igual).
# Retorna una hashtable [clau] -> registre.
function _FlattenInformesDb($db) {
    $map = @{}
    if ($null -eq $db -or $null -eq $db.activitats) { return $map }
    $arrel = [string](_PropInf $db 'carpeta_arrel')
    foreach ($act in $db.activitats) {
        if ($null -eq $act.informes) { continue }
        foreach ($inf in $act.informes) {
            $ruta = [string]$inf.ruta
            if ([string]::IsNullOrWhiteSpace($ruta)) { continue }
            $teMarca = ($null -ne $inf.PSObject.Properties['editat_a_ma'])
            _InferEditatAMa $inf
            $motiuStr = if ($null -ne $inf.PSObject.Properties['motiu']) { [string]$inf.motiu } else { '' }
            $motius = if ([string]::IsNullOrWhiteSpace($motiuStr)) { @() } else { @($motiuStr -split ',\s*') }
            $ign = $false
            if ($null -ne $inf.PSObject.Properties['ignorat']) { $ign = [bool]$inf.ignorat }
            $conclusioText = [string]$inf.conclusio
            # Compatibilitat: si la base es d'abans d'aquest camp, la calculem
            # ara mateix (no cal reescanejar per tenir-la la primera vegada).
            $conclusioBreu = if ($null -ne $inf.PSObject.Properties['conclusio_breu']) { [string]$inf.conclusio_breu } else { _ConclusioBreu $conclusioText }
            $map[(_ClauInforme $ruta $arrel)] = [pscustomobject]@{
                Data          = [string]$inf.data
                Gia           = [string]$act.id_gia
                GiaFont       = [string](_PropInf $inf 'gia_font')
                Expedient     = [string]$act.expedient
                Titular       = [string]$act.titular
                Conclusio     = $conclusioText
                ConclusioBreu = $conclusioBreu
                Tipus         = [string](_PropInf $inf 'tipus')
                Fitxer        = [string]$inf.fitxer
                Ruta          = $ruta
                Carpeta       = [string]$act.carpeta
                Modificat     = if ($null -ne $inf.PSObject.Properties['modificat']) { [string]$inf.modificat } else { '' }
                Ignorat       = $ign
                Motius        = $motius
                EditatAMa         = [bool](_PropInf $inf 'editat_a_ma')
                TeMarcaEdicio     = $teMarca
                AutoConclusioBreu = [string](_PropInf $inf 'auto_conclusio_breu')
                AutoIgnorat       = [bool](_PropInf $inf 'auto_ignorat')
            }
        }
    }
    return $map
}

# En reprocessar un informe que ja era a la base: el de l'usuari PREVAL si l'hi
# havia corregit; si no, mana el que acaba de sortir de l'informe (que potser
# s'ha tornat a escriure). $prev: el registre de _FlattenInformesDb; $nou: el
# de Get-InformeData. Retorna $nou, amb les correccions si n'hi ha. PURA.
#
# Una base d'abans de la marca (el registre no la porta) es tracta com
# _InferEditatAMa, i a mes, aqui si que es pot dir de l'ignorat: si el de la
# base no es el per defecte que acaba de sortir, l'havia posat l'usuari. (Des
# de l'octubre de 2026 el per defecte es sempre "no ignorat": en una base
# d'abans de la marca, una MNS que s'ignorava per defecte es llegira com a
# ignorada a ma. Les bases d'ara ja porten la marca a tots els informes.)
function _AplicaEdicioPrevia($prev, $nou) {
    $manual = [bool]$prev.EditatAMa
    if (-not [bool]$prev.TeMarcaEdicio) {
        $manual = $manual -or ([bool]$prev.Ignorat -ne [bool]$nou.Ignorat)
    }
    if (-not $manual) { return $nou }
    # L'automatic ja diu el mateix que la correccio: ja no ho es (179 de les
    # "correccions" de la base d'octubre de 2026 eren defectes del classificador,
    # no gustos de l'usuari, i pintaven mitja base en vermell a l'editor). Es
    # treu la marca; la que no coincideix, es queda.
    if ([string]$prev.ConclusioBreu -eq [string]$nou.ConclusioBreu -and [bool]$prev.Ignorat -eq [bool]$nou.Ignorat) { return $nou }
    $nou.AutoConclusioBreu = [string]$nou.ConclusioBreu
    $nou.AutoIgnorat = [bool]$nou.Ignorat
    $nou.ConclusioBreu = [string]$prev.ConclusioBreu
    $nou.Ignorat = [bool]$prev.Ignorat
    $nou.EditatAMa = $true
    return $nou
}

# Un informe tal com es desa a informes-db.json. Els camps de l'edicio a ma
# (auto_*) nomes si ho esta: la base no s'omple de camps buits. PURA.
function _InformeAJson($r) {
    $o = [pscustomobject]@{
        data           = $r.Data
        fitxer         = $r.Fitxer
        ruta           = $r.Ruta
        conclusio      = $r.Conclusio
        conclusio_breu = $r.ConclusioBreu
        tipus          = [string]$r.Tipus
        modificat      = $r.Modificat
        ignorat        = [bool]$r.Ignorat
        motiu          = (@($r.Motius) -join ', ')
        editat_a_ma    = [bool]$r.EditatAMa
    }
    # D'on surt l'ID GIA ('document', 'assignat', 'carpeta', 'excel', 'germans'),
    # nomes si se sap: les bases d'abans no ho portaven.
    $giaFont = [string](_PropInf $r 'GiaFont')
    if ($giaFont -ne '') { Add-Member -InputObject $o -NotePropertyName gia_font -NotePropertyValue $giaFont }
    if ([bool]$r.EditatAMa) {
        Add-Member -InputObject $o -NotePropertyName auto_conclusio_breu -NotePropertyValue ([string]$r.AutoConclusioBreu)
        Add-Member -InputObject $o -NotePropertyName auto_ignorat -NotePropertyValue ([bool]$r.AutoIgnorat)
    }
    return $o
}

# ----------------------------------------------------------------------------
# Peces PURES de l'escaneig (les fan servir el nucli i ValidarClassificacio.ps1)
# ----------------------------------------------------------------------------

# La carpeta on viu un informe (la ruta sense el nom del fitxer), per comparar-la
# amb la dels altres. Parteix per les dues barres (es prova a Linux).
function _DirInforme([string]$ruta) {
    $i = [Math]::Max($ruta.LastIndexOf('\'), $ruta.LastIndexOf('/'))
    if ($i -lt 0) { return '' }
    return $ruta.Substring(0, $i)
}

# Un informe SENSE ID GIA en una carpeta on TOTS els altres que en tenen son del
# mateix GIA, va amb aquell GIA. Cas real: un informe de llicencia sense GIA a la
# capcalera, en la carpeta del seu expedient (que no diu "GIA n" al nom),
# quedava com una activitat a part. _GiaFromFolderName nomes mira el NOM de la
# carpeta. Toca els registres ($informes: Gia, GiaFont, Ruta, Motius) i torna
# quants n'ha resolt.
function _GiaDelsGermans($informes) {
    $perDir = @{}
    foreach ($r in @($informes)) {
        if ($null -eq $r -or [string]::IsNullOrWhiteSpace([string]$r.Gia)) { continue }
        $d = _DirInforme ([string]$r.Ruta)
        if (-not $perDir.ContainsKey($d)) { $perDir[$d] = @{} }
        $perDir[$d][[string]$r.Gia] = $true
    }
    $n = 0
    foreach ($r in @($informes)) {
        if ($null -eq $r -or -not [string]::IsNullOrWhiteSpace([string]$r.Gia)) { continue }
        $d = _DirInforme ([string]$r.Ruta)
        if (-not $perDir.ContainsKey($d) -or $perDir[$d].Count -ne 1) { continue }
        $r.Gia = [string]@($perDir[$d].Keys)[0]
        $r.GiaFont = 'germans'
        $r.Motius = @(@($r.Motius) | Where-Object { $_ -ne 'sense ID GIA' })
        $n++
    }
    return $n
}

# Agrupa els registres per activitat: per ID GIA quan n'hi ha; si NO en tenen,
# per CARPETA (tots els informes d'una mateixa carpeta = una activitat). Torna
# les activitats tal com es desen (informes per ordre, _OrdenaInformesActivitat,
# i l'estat calculat), per ID GIA NUMERIC (com a text, '10' anava abans que '9').
function _AgrupaInformes($informes) {
    $groups = [ordered]@{}
    foreach ($r in @($informes)) {
        if ($null -eq $r) { continue }
        $key = if (-not [string]::IsNullOrWhiteSpace($r.Gia)) { "GIA:$($r.Gia)" }
               else { "DIR:$($r.Carpeta)" }
        if (-not $groups.Contains($key)) {
            $groups[$key] = [pscustomobject]@{
                id_gia    = $r.Gia
                expedient = $r.Expedient
                titular   = $r.Titular
                carpeta   = $r.Carpeta
                _informes = (New-Object System.Collections.ArrayList)
            }
        }
        $g = $groups[$key]
        # Emplenem camps de l'activitat si encara estan buits.
        if ([string]::IsNullOrWhiteSpace($g.expedient) -and -not [string]::IsNullOrWhiteSpace($r.Expedient)) { $g.expedient = $r.Expedient }
        if ([string]::IsNullOrWhiteSpace($g.titular)   -and -not [string]::IsNullOrWhiteSpace($r.Titular))   { $g.titular = $r.Titular }
        [void]$g._informes.Add((_InformeAJson $r))
    }
    $activitats = New-Object System.Collections.ArrayList
    foreach ($g in $groups.Values) {
        $act = [pscustomobject]@{
            id_gia       = $g.id_gia
            expedient    = $g.expedient
            titular      = $g.titular
            carpeta      = $g.carpeta
            estat_actual = ''
            informes     = @(_OrdenaInformesActivitat $g._informes)
        }
        $act.estat_actual = _EstatActualActivitat $act
        [void]$activitats.Add($act)
    }
    return @($activitats | Sort-Object { _GiaNumeric $_.id_gia }, { [string]$_.carpeta })
}

# Mateixa carpeta d'informes? (sense distingir majuscules ni la barra final)
function _MateixaArrel([string]$a, [string]$b) {
    return (($a -replace '/', '\').TrimEnd('\') -ieq ($b -replace '/', '\').TrimEnd('\'))
}

# La base es d'una ALTRA carpeta d'informes i gairebe no hi casa res: escanejar
# la reescriuria sencera i s'hi perdrien les correccions a ma dels informes que
# no casen. Torna el text de la pregunta, o '' si no cal preguntar (la mateixa
# arrel, casa la majoria, o no es perdria cap correccio). PURA.
function _AvisCanviArrel([string]$arrelBase, [string]$arrelAra, [int]$nBase, [int]$nCasats, [int]$nCorrPerdudes) {
    if ($nBase -le 0 -or [string]::IsNullOrWhiteSpace($arrelBase) -or (_MateixaArrel $arrelBase $arrelAra)) { return '' }
    if ($nCasats * 2 -ge $nBase -or $nCorrPerdudes -le 0) { return '' }
    return ("La base d'informes " + [char]0x00E9 + "s de:`n" + $arrelBase + "`n`nAra s'escanejaria:`n" + $arrelAra +
            "`n`nNom" + [char]0x00E9 + "s hi casen " + $nCasats + " dels " + $nBase + " informes de la base: s'hi perdrien " +
            $nCorrPerdudes + " correccions fetes a m" + [char]0x00E0 + ".`n`nVols continuar igualment?")
}

# ----------------------------------------------------------------------------
# EL NUCLI de l'escaneig, SENSE CAP FINESTRA: el fan servir el boto (amb la
# seva finestra de progres) i el mode automatic (en segon pla, sense res).
# $onProgres (opcional): & $onProgres <text> <fets> <total> ($total 0 = encara no
# se sap). $onConfirma (opcional): & $onConfirma <pregunta> -> $true per seguir;
# sense ell (l'automatic) la resposta es NO. Torna @{ Ok; Error; NInformes;
# Reprocessats; NActivitats; NRevisar; OutPath } (Cancelat = $true si s'ha dit
# que no). Si alguna cosa peta a mitges, llanca (i no s'ha escrit res).
# ----------------------------------------------------------------------------
function _InformesDirAccessible([string]$dir) {
    if ([string]::IsNullOrWhiteSpace($dir)) { return $false }
    try { return [bool](Test-Path -LiteralPath $dir -ErrorAction SilentlyContinue) } catch { return $false }
}

# $giaAssignatsPath: el fitxer d'assignacions de GIA; buit = el mes recent de
# $LocalActivitatsDir (ValidarClassificacio el busca abans de canviar la carpeta
# local per una de temporal i el passa aqui).
function Invoke-InformesDbEscaneig([scriptblock]$onProgres = $null, [scriptblock]$onConfirma = $null, [string]$giaAssignatsPath = '') {
    # La carpeta d'informes. Si la unitat (la I: de la feina) no hi es, no es un
    # error del programa: potser s'esta fora de la feina.
    $dir = $InformesDir
    if (-not (_InformesDirAccessible $dir)) {
        return @{ Ok = $false; Error = "No s'ha trobat la carpeta d'informes: $dir" }
    }
    $avisa = { param($t, $i, $n) if ($null -ne $onProgres) { & $onProgres $t $i $n } }

    # 3. Carregar l'Excel d'activitats (opcional; per la cerca inversa i el
    #    titular). Si no hi ha Excel, es continua sense aquest fallback.
    $cache = $null; $expToGia = $null
    try {
        $excel = Find-LatestActivitatsExcel
        if ($null -ne $excel) {
            & $avisa "Llegint la base d'activitats (Excel)..." 0 0
            $cache = Initialize-ActivitatsCache $excel.File
            $expToGia = Build-ExpedientToGiaMap $cache
        }
    } catch { $cache = $null; $expToGia = $null }

    # 3b. Carregar la base anterior (si existeix) per fer un escaneig
    #     INCREMENTAL: nomes es reobren els .docx modificats DESPRES de
    #     l'ultima actualitzacio; la resta es reutilitzen. Els fitxers que ja
    #     no existeixen es podaran sols (nomes reagrupem els que trobem ara).
    #     Si no hi ha base previa (o esta corrupta), es fa un escaneig complet.
    #     Si la base es d'una altra versio del classificador, es tornen a llegir
    #     TOTS (vegeu $Script:ClassificadorVersio), pero es conserven les
    #     correccions a ma.
    # 3a. L'ID GIA assignat a ma (vegeu Read-GiaAssignats). Si el fitxer ha
    #     canviat des de l'ultima vegada, es tornen a llegir TOTS els informes.
    $assignats = Read-GiaAssignats $giaAssignatsPath
    $outPath    = Get-InformesDbPath
    $prevByRuta = @{}
    $prevUtc    = [datetime]::MinValue
    $prevArrel  = ''
    $generatEl  = (Get-Date).ToString('o')
    if (Test-Path -LiteralPath $outPath) {
        try {
            $prevDb     = Read-JsonFile $outPath
            # Una base sense carpeta_arrel (no n'hi hauria d'haver cap) es
            # dona per feta amb la carpeta d'ara: si no, cap informe hi
            # casaria per la ruta relativa i es perdrien les correccions.
            if ([string]::IsNullOrWhiteSpace([string](_PropInf $prevDb 'carpeta_arrel'))) {
                Add-Member -InputObject $prevDb -NotePropertyName carpeta_arrel -NotePropertyValue $dir -Force
            }
            $prevByRuta = _FlattenInformesDb $prevDb
            $prevArrel  = [string](_PropInf $prevDb 'carpeta_arrel')
            if ($prevDb.PSObject.Properties['actualitzat_el'] -and -not [string]::IsNullOrWhiteSpace([string]$prevDb.actualitzat_el)) {
                try { $prevUtc = ([datetime]::Parse([string]$prevDb.actualitzat_el)).ToUniversalTime() } catch { $prevUtc = [datetime]::MinValue }
            }
            if ([string](_PropInf $prevDb 'versio_classificador') -ne $Script:ClassificadorVersio) { $prevUtc = [datetime]::MinValue }
            if ([string](_PropInf $prevDb 'gia_assignats') -ne [string]$assignats.Signatura) { $prevUtc = [datetime]::MinValue }
            if ($prevDb.PSObject.Properties['generat_el'] -and -not [string]::IsNullOrWhiteSpace([string]$prevDb.generat_el)) {
                $generatEl = [string]$prevDb.generat_el
            }
        } catch { $prevByRuta = @{}; $prevUtc = [datetime]::MinValue; $prevArrel = '' }
    }

    # 4. Recollir els fitxers candidats (.docx o .doc amb data al principi
    #    del nom), NOMES del primer nivell (Get-FitxersPrimerNivell). Sense
    #    -Filter i filtrem per extensio nosaltres: evita el parany de "*.doc"
    #    -Filter que a vegades tambe encerta ".docx" pel nom curt (8.3) de NTFS.
    & $avisa "Cercant informes a:`n$dir" 0 0
    # Primer s'ASSIGNA i despres es filtra: Get-FitxersPrimerNivell torna la
    # llista amb coma, i passada directament per la canonada arribava al
    # Where-Object com UN SOL objecte (la llista sencera): el filtre deixava
    # passar TOTS els fitxers en cop que n'hi hagues un de bo (mesurat).
    $primerNivell = Get-FitxersPrimerNivell $dir
    $allInformes = $primerNivell |
                   Where-Object {
                       $_.Name -notlike '~$*' -and
                       ($_.Extension -ieq '.docx' -or $_.Extension -ieq '.doc') -and
                       $null -ne (_ParseDataInformeFromName $_.Name)
                   }
    $files = @($allInformes)
    $total = $files.Count

    # 4b. La base es d'una altra carpeta i no hi casa gairebe res: PREGUNTAR
    #     abans d'escriure (l'automatic, que no pot preguntar, no escriu).
    $claus = @{}
    foreach ($f in $files) { $claus[(_ClauInforme $f.FullName $dir)] = $true }
    $nCasats = 0; $nCorrPerdudes = 0
    foreach ($kv in $prevByRuta.GetEnumerator()) {
        if ($claus.ContainsKey($kv.Key)) { $nCasats++ } elseif ([bool]$kv.Value.EditatAMa) { $nCorrPerdudes++ }
    }
    $pregunta = _AvisCanviArrel $prevArrel $dir $prevByRuta.Count $nCasats $nCorrPerdudes
    if ($pregunta -ne '' -and ($null -eq $onConfirma -or -not (& $onConfirma $pregunta))) {
        return @{ Ok = $false; Cancelat = $true; Error = ("No s'ha actualitzat la base. " + ($pregunta -replace "`n`nVols continuar igualment\?$", '')) }
    }

    # 5. Analitzar cada informe (incremental: reutilitzem els no modificats).
    #    Word només es crea (mandrosament) si cal reprocessar algun .doc
    #    antic; es tanca sempre al 'finally', encara que hi hagi un error.
    $informes = New-Object System.Collections.ArrayList
    $reprocessats = 0
    $i = 0
    $wordApp = $null
    try {
        foreach ($f in $files) {
            $i++
            $clau = _ClauInforme $f.FullName $dir
            $teEntrada = $prevByRuta.ContainsKey($clau)
            if (-not (_HaDeReprocessar $f.LastWriteTimeUtc $prevUtc $teEntrada)) {
                # No s'ha tocat des de l'ultim escaneig: reutilitzem l'entrada
                # (amb la ruta d'ara: la base pot ser d'una altra unitat).
                $r = $prevByRuta[$clau]
                $r.Ruta = $f.FullName
            } else {
                if ($f.Extension -ieq '.doc' -and $null -eq $wordApp) {
                    # -Opcional: sense Word, _ReadInformeParagraphs torna @()
                    # i l'informe es queda sense conclusio, pero l'escaneig
                    # continua. New-WordApp hi afegeix l'AutomationSecurity,
                    # que aqui compta: aquests .doc son a la unitat de xarxa.
                    $wordApp = New-WordApp -Opcional
                }
                $r = Get-InformeData $f $expToGia $cache $wordApp $assignats.Mapa[$clau]
                # El que l'usuari hagi corregit a ma ("Editar base") PREVAL;
                # la resta, la mana el que acaba de sortir de l'informe
                # (vegeu _AplicaEdicioPrevia).
                if ($teEntrada) { $r = _AplicaEdicioPrevia $prevByRuta[$clau] $r }
                $reprocessats++
            }
            if (($i % 5) -eq 0 -or $i -eq $total) {
                & $avisa "Analitzant informes... ($i de $total, $reprocessats de nous/modificats)" $i $total
            }
            [void]$informes.Add($r)
        }
    } finally {
        if ($null -ne $wordApp) { try { $wordApp.Quit() } catch { } }
    }

    # 6. Els que no tenen GIA, el dels germans de carpeta; agrupar per
    #    activitat; i el que s'ha de revisar (DESPRES dels germans: un GIA
    #    resolt ja no ho es).
    [void](_GiaDelsGermans $informes)
    $activitatsOrd = @(_AgrupaInformes $informes)
    $revisar = New-Object System.Collections.ArrayList
    foreach ($r in $informes) {
        if (@($r.Motius).Count -gt 0) {
            [void]$revisar.Add([pscustomobject]@{
                fitxer = $r.Fitxer
                ruta   = $r.Ruta
                motiu  = (@($r.Motius) -join ', ')
            })
        }
    }

    # 7. Escriure el JSON (conservem generat_el; actualitzat_el = ara).
    $outObj = [pscustomobject]@{
        generat_el           = $generatEl
        actualitzat_el       = (Get-Date).ToString('o')
        versio_classificador = $Script:ClassificadorVersio
        gia_assignats        = [string]$assignats.Signatura
        carpeta_arrel        = $dir
        n_informes           = $informes.Count
        n_activitats         = $activitatsOrd.Count
        activitats           = $activitatsOrd
        a_revisar            = @($revisar)
    }
    Write-JsonFile $outPath $outObj 8

    # Les assignacions que ja no troben l'informe (l'han mogut o reanomenat): no
    # peta, es diu (el resum del boto, el registre de l'automatic i
    # ValidarClassificacio).
    $noTrobats = @(_GiaAssignatsNoTrobats $assignats $claus)

    return @{ Ok = $true; Error = ''; NInformes = $informes.Count; Reprocessats = $reprocessats
              NActivitats = $activitatsOrd.Count; NRevisar = $revisar.Count; OutPath = $outPath
              GiaAssignatsFitxer = [string]$assignats.Fitxer; GiaAssignatsError = [string]$assignats.Error
              GiaAssignatsNoTrobats = $noTrobats }
}

# Nom del mutex de l'escaneig: el comparteixen el boto i l'automatic, perque dos
# escaneigs alhora escriurien informes-db.json l'un sobre l'altre.
$Script:BaseMutexNom = 'Global\InformesCornella.BaseInformes'

# ----------------------------------------------------------------------------
# El BOTO "Actualitzar base": el nucli amb finestra de progres i un resum.
# ----------------------------------------------------------------------------
function Invoke-InformesDbScan {
    $dir = $InformesDir
    if (-not (_InformesDirAccessible $dir)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat la carpeta d'informes:`n$dir`n`nSi treballes fora de la feina (sense la unitat I:), obre-la quan hi tinguis accés. Pots canviar la ruta a Configuració.",
            'Base d''informes', 'OK', 'Warning') | Out-Null
        return
    }

    $form = _NewForm
    $form.Text = "Actualitzant base d'informes"
    $form.Size = New-Object System.Drawing.Size(560, 170)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 20)
    $lbl.Size = New-Object System.Drawing.Size(510, 60)
    $lbl.Text = "Cercant informes a:`n$dir"
    $form.Controls.Add($lbl)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 90)
    $bar.Size = New-Object System.Drawing.Size(510, 24)
    $bar.Style = 'Marquee'
    $form.Controls.Add($bar)
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $onProgres = {
        param($text, $fets, $total)
        $lbl.Text = $text
        if ($total -gt 0) {
            if ($bar.Style -ne 'Continuous') { $bar.Style = 'Continuous'; $bar.Minimum = 0 }
            $bar.Maximum = [Math]::Max(1, $total)
            $bar.Value = [Math]::Min($bar.Maximum, $fets)
        }
        [System.Windows.Forms.Application]::DoEvents()
    }.GetNewClosure()
    # La base es d'una altra carpeta d'informes (vegeu _AvisCanviArrel).
    $onConfirma = {
        param($pregunta)
        $resp = [System.Windows.Forms.MessageBox]::Show($pregunta, 'Base d''informes', 'YesNo', 'Warning', 'Button2')
        return ($resp -eq [System.Windows.Forms.DialogResult]::Yes)
    }

    $caixa = @{ Res = $null }
    $res = $null
    $err = ''
    try {
        # Si el mode automatic esta escanejant ara mateix, no se'n fa un altre
        # al damunt (escriurien la base l'un sobre l'altre).
        $fet = Invoke-AmbMutexUnic $Script:BaseMutexNom {
            $caixa.Res = Invoke-InformesDbEscaneig $onProgres $onConfirma
        }
        if ($fet) { $res = $caixa.Res } else { $err = 'ocupat' }
    } catch {
        $err = $_.Exception.Message
    } finally {
        try { $form.Close() } catch { }
    }
    if ($err -eq 'ocupat') {
        [System.Windows.Forms.MessageBox]::Show(
            "La base d'informes s'est" + [char]0x00E0 + " actualitzant sola ara mateix (mode autom" + [char]0x00E0 + "tic).`n`nTorna-hi d'aqu" + [char]0x00ED + " a una estona.",
            'Base d''informes', 'OK', 'Information') | Out-Null
        return
    }
    if ($err -ne '' -or $null -eq $res) {
        [System.Windows.Forms.MessageBox]::Show("Error escanejant els informes:`n$err", 'Base d''informes', 'OK', 'Error') | Out-Null
        return
    }
    if (-not $res.Ok) {
        if ([bool]$res.Cancelat) { return }
        [System.Windows.Forms.MessageBox]::Show([string]$res.Error, 'Base d''informes', 'OK', 'Warning') | Out-Null
        return
    }
    # L'ultima l'has feta tu: la data del menu deixa de sortir en verd.
    [void](_BaseAutoDesaEstat @{ mode = 'manual' })
    $msg = "Base d'informes actualitzada.`n`n" +
           "Informes trobats: $($res.NInformes)`n" +
           "Nous o modificats (reprocessats): $($res.Reprocessats)`n" +
           "Activitats: $($res.NActivitats)`n" +
           "A revisar: $($res.NRevisar)`n" +
           $(if ([string]$res.GiaAssignatsError) { "`nATENCI" + [char]0x00D3 + ": " + [string]$res.GiaAssignatsError + "`n" } else { '' }) +
           $(if (@($res.GiaAssignatsNoTrobats).Count -gt 0) { "`nAssignacions de GIA que ja no troben l'informe: " + @($res.GiaAssignatsNoTrobats).Count + " (les llista ValidarClassificacio)`n" } else { '' }) +
           "`n" +
           "Fitxer:`n$($res.OutPath)`n`nVols obrir-lo?"
    $r = [System.Windows.Forms.MessageBox]::Show($msg, 'Base d''informes', 'YesNo', 'Information')
    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
        try { Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$($res.OutPath)`"" | Out-Null } catch { }
    }
}

# ============================================================================
# MODE AUTOMATIC (l'interruptor A/M de sota la rajola "Actualitzar base")
# ============================================================================
# L'usuari: "ja que sera tan important per fer el Planol activitats, que tambe
# tingui l'opcio d'actualitzar-se automaticament (igual que Copiar informes)".
# Mateixa regla i mateixa hora (ModeAutomatic.ps1, $Script:AutoHora: les 13:00)
# amb el programa obert i, si l'ultima vegada que tocava no es va fer, en obrir
# el programa;
# en un proces a part (BaseInformesAuto.ps1), sense res a la pantalla.
#
# Les correccions a ma ("Editar base") PREVALEN igual que amb el boto: la
# passada es el mateix Invoke-InformesDbEscaneig.
#
# L'ESTAT, a informes-db-auto.json (no dins de la base: l'editor la reescriu
# sencera i s'hi perdria):
#   auto     l'interruptor
#   auto_el  l'ultima PASSADA automatica, encara que no hagi pogut fer res: es
#            la marca que diu que el venciment ja s'ha servit
#   mode     'auto' | 'manual': qui va fer l'ultima actualitzacio
$Script:BaseEstatPlantilla = [ordered]@{ auto = $false; auto_el = ''; mode = '' }

function _BaseAutoStatePath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'informes-db-auto.json')
}

function _BaseAutoEstat { return (Read-EstatAuto (_BaseAutoStatePath) $Script:BaseEstatPlantilla) }

function _BaseAutoDesaEstat($canvis) { return (Save-EstatAuto (_BaseAutoStatePath) $canvis $Script:BaseEstatPlantilla) }

function _BaseAutoToca([datetime]$ara, $ultimAuto) {
    return (Test-ProgramacioToca 'informesdb' $ara $ultimAuto)
}

function _BaseAutoLog([string]$msg) { Write-AutoLog 'informes-db-log.txt' $msg }

# La passada (la crida BaseInformesAuto.ps1). Cap finestra ni cap pregunta: tot
# va al registre. Apunta 'auto_el' SEMPRE (si no, el menu la tornaria a llancar
# cada minut), tambe quan la carpeta no hi es o quan el boto esta escanejant
# ara mateix (llavors la base ja s'esta posant al dia). Torna el resultat de
# l'escaneig, o @{ Ok = $false; Error }.
function Invoke-InformesDbAuto {
    $ini = @{ auto_el = (Get-Date).ToString('o') }
    $caixa = @{ Res = $null; Err = '' }
    $lliure = $false
    try {
        $lliure = Invoke-AmbMutexUnic $Script:BaseMutexNom {
            $caixa.Res = Invoke-InformesDbEscaneig $null
        }
    } catch { $caixa.Err = [string]$_.Exception.Message }
    if (-not $lliure -and $caixa.Err -eq '') {
        [void](_BaseAutoDesaEstat $ini)
        _BaseAutoLog 'Passada automatica: ja s''esta actualitzant, no es fa res.'
        return @{ Ok = $false; Error = 'ocupat' }
    }
    if ($caixa.Err -ne '') {
        [void](_BaseAutoDesaEstat $ini)
        _BaseAutoLog ('ERROR: ' + $caixa.Err)
        return @{ Ok = $false; Error = $caixa.Err }
    }
    $res = $caixa.Res
    if (-not [bool]$res.Ok) {
        [void](_BaseAutoDesaEstat $ini)
        _BaseAutoLog ('ATURAT: ' + ([string]$res.Error -replace "`r?`n", ' '))
        return $res
    }
    [void](_BaseAutoDesaEstat @{ auto_el = $ini['auto_el']; mode = 'auto' })
    _BaseAutoLog ("Passada automatica: informes=$($res.NInformes) reprocessats=$($res.Reprocessats) " +
                  "activitats=$($res.NActivitats) a_revisar=$($res.NRevisar)" +
                  $(if ([string]$res.GiaAssignatsError) { ' AVIS: ' + [string]$res.GiaAssignatsError } else { '' }) +
                  $(if (@($res.GiaAssignatsNoTrobats).Count -gt 0) { ' assignacions_GIA_no_trobades=' + @($res.GiaAssignatsNoTrobats).Count } else { '' }))
    return $res
}

function Start-InformesDbAuto { return (Start-ProcesAutoUnic 'base' 'BaseInformesAuto.ps1') }

# El menu ho crida en obrir-se i a cada minut.
function Invoke-BaseAutoSiToca {
    $est = _BaseAutoEstat
    if (-not [bool]$est['auto']) { return $false }
    if (-not (_BaseAutoToca (Get-Date) $est['auto_el'])) { return $false }
    return (Start-InformesDbAuto)
}

Register-ProgramacioAuto 'informesdb' 'Actualitzar base'
$Script:ModesAuto['informesdb'] = @{
    Titol     = 'Actualitzar base'
    Actiu     = { $e = _BaseAutoEstat; [bool]$e['auto'] }
    DesaActiu = { param($on) _BaseAutoDesaEstat @{ auto = [bool]$on } }
    UltimMode = { $e = _BaseAutoEstat; [string]$e['mode'] }
    SiToca    = { Invoke-BaseAutoSiToca }
    Requisit  = {
        if (-not [string]::IsNullOrWhiteSpace($InformesDir)) { return '' }
        return ("Per actualitzar la base sola cal dir on s" + [char]0x00F3 + "n els informes.`n`nVes a Configuraci" + [char]0x00F3 + " (el bot" + [char]0x00F3 + " de la roda, a dalt a la dreta) i indica la carpeta d'informes.")
    }
    TipA      = { Get-AutoTipText "la base s'actualitza sola" 'informesdb' }
    TipM      = "Mode MANUAL: nomes s'actualitza quan cliques la rajola. Clica per posar-ho en automatic."
}
