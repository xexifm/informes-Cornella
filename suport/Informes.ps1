#requires -Version 5.1
<#
.SYNOPSIS
  Escaner d'informes ja generats -> base de dades JSON.

.DESCRIPTION
  Recorre l'arbre de carpetes dels informes (per defecte $InformesDir, germa de
  la carpeta de l'Excel d'activitats) i, per cada informe (.docx, o .doc antic
  via Word COM), en treu:
    - la DATA (del principi del nom del fitxer),
    - l'ID GIA (del document -ignorant placeholders com "-"/"XXX" quan encara
      no n'hi ha-; si no hi es, del nom de la carpeta "GIA 361"; si tampoc, de
      l'Excel d'activitats cercant per numero d'expedient),
    - la CONCLUSIO (el paragraf que comenca amb una de les frases de
      $Script:ConclusioStartPhrases: "Vist l'anterior" i "Tenint en
      consideracio el risc" son fiables; "S'informa favorablement" i "El
      titular/L'organitzador es responsable d'executar" es desen igualment,
      pero l'informe queda marcat "ignorat" PER DEFECTE (nomes la primera
      vegada; l'usuari pot desmarcar-ho des de l'editor) perque son clausules
      molt semblants entre informes diferents.
  Ho desa AGRUPAT PER ACTIVITAT a local\base-dades-activitats\informes-db.json
  (carpeta ignorada per git). Els informes que no es poden resoldre del tot
  van a un bloc "a_revisar".

  Es un modul del motor: es carrega (dot-source) des de GenerarInforme.ps1, aixi
  reutilitza les funcions de lectura de .docx sense Word (de Seguiment.ps1),
  _NormalitzaText i l'acces a l'Excel d'activitats (Find-LatestActivitatsExcel /
  Initialize-ActivitatsCache). Les funcions de logica de text son PURES (operen
  sobre cadenes) perque es puguin provar en headless (Linux, sense Word); la
  lectura de .doc antics (Word COM) nomes es prova manualment a Windows.

.NOTES
  Es llanca des del menu (Pas 1) amb el boto "Actualitzar base d'informes".
#>

# ----------------------------------------------------------------------------
# Logica de text (funcions PURES, testejables en headless)
# ----------------------------------------------------------------------------

# Treu la data del PRINCIPI del nom del fitxer. Tolerant amb el format:
#   AAAA-MM-DD / AAAAMMDD / AAAA.MM.DD / AAAA_MM_DD  (any de 4 xifres)
#   AA-MM-DD   / AA.MM.DD / AA_MM_DD                 (any de 2 xifres, amb
#                                                     separador obligatori per no
#                                                     confondre-ho amb AAAAMMDD)
# Retorna 'yyyy-MM-dd' o $null si el nom no comenca amb una data valida.
# (El programa considera "informe" NOMES els fitxers que comencen amb data.)
function _ParseDataInformeFromName($name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return $null }
    $n = [string]$name
    # Any de 4 xifres, separadors opcionals.
    if ($n -match '^(\d{4})[-_.]?(\d{2})[-_.]?(\d{2})(?:\D|$)') {
        $y = [int]$Matches[1]; $mo = [int]$Matches[2]; $d = [int]$Matches[3]
        if ($y -ge 1990 -and $y -le 2100 -and $mo -ge 1 -and $mo -le 12 -and $d -ge 1 -and $d -le 31) {
            return ('{0:D4}-{1:D2}-{2:D2}' -f $y, $mo, $d)
        }
    }
    # Any de 2 xifres, separador OBLIGATORI.
    if ($n -match '^(\d{2})[-_.](\d{2})[-_.](\d{2})(?:\D|$)') {
        $y = 2000 + [int]$Matches[1]; $mo = [int]$Matches[2]; $d = [int]$Matches[3]
        if ($mo -ge 1 -and $mo -le 12 -and $d -ge 1 -and $d -le 31) {
            return ('{0:D4}-{1:D2}-{2:D2}' -f $y, $mo, $d)
        }
    }
    return $null
}

# Cert si el valor trobat despres de "ID GIA:" es un placeholder (activitat
# encara sense GIA assignat: "-", "XXX", "N/A"...) i no un ID real. Cal
# distingir-ho perque, si no, activitats totalment diferents que encara no
# tenen GIA queden ajuntades sota una mateixa "activitat" fantasma amb
# id_gia="-" (vist a la carpeta real d'informes).
function _EsGiaPlaceholder($val) {
    if ([string]::IsNullOrWhiteSpace($val)) { return $true }
    $n = $val.Trim().ToLowerInvariant()
    return ($n -eq '-' -or $n -eq '--' -or $n -eq '---' -or $n -eq 'xxx' -or $n -eq 'n/a' -or $n -eq 'na' -or $n -eq '?')
}

# ID GIA d'una llista de linies de text del document. Busca la linia que conte
# "ID GIA" i en retorna el valor (despres dels dos punts / espais). '' si no hi
# es, o si el valor trobat es un placeholder de "encara sense GIA".
function _ExtractIdGia($lines) {
    foreach ($ln in $lines) {
        $s = [string]$ln
        if ($s -imatch 'ID\s*GIA\s*:?\s*(.+)$') {
            $val = $Matches[1].Trim()
            # Ens quedem nomes amb el primer "token" del valor (l'ID; evita
            # arrossegar text si la linia porta res mes al darrere).
            $token = $val
            if ($val -match '^([\w./-]+)') { $token = $Matches[1] }
            if (_EsGiaPlaceholder $token) { continue }
            return $token
        }
    }
    return ''
}

# Numero d'expedient d'una llista de linies. Busca la linia que comenca per
# "Exp" (p.ex. "Exp. Num: 2025/1/2563") i retorna el valor despres dels dos punts.
function _ExtractExpedient($lines) {
    foreach ($ln in $lines) {
        $s = [string]$ln
        $nt = _NormalitzaText $s
        if ($nt -match '^exp') {
            $idx = $s.IndexOf(':')
            if ($idx -ge 0) { return $s.Substring($idx + 1).Trim() }
            if ($s -match '(\S+/\S+/\S+)') { return $Matches[1] }
        }
    }
    return ''
}

# Normalitzacio per comparar frases: sense accents, minuscules i SENSE
# apostrofs. Els informes fan servir l'apostrof TIPOGRAFIC (U+2019), pero les
# nostres frases de referencia el recte (U+0027); traient-los tots dos (i altres
# variants) la comparacio casa igual. Fem servir codepoints [char]0x.... per no
# dependre de l'encoding amb que PowerShell 5.1 llegeix aquest fitxer.
function _ConclNorm($s) {
    $t = _NormalitzaText $s
    $apos = @([char]0x0027, [char]0x2018, [char]0x2019, [char]0x02BC, [char]0x00B4, [char]0x0060)
    foreach ($a in $apos) { $t = $t.Replace([string]$a, '') }
    return $t
}

# Frases que poden marcar l'INICI de la conclusio d'un informe: cada familia
# de tramits tanca la decisio d'una manera diferent (vist a la carpeta real
# d'informes). Font: 'vist_anterior' i 'risc' son fiables (frase de decisio
# propia i diferenciada de cada informe, no repetida literalment d'un informe
# a l'altre); 'mns' i 'act_extr' es desen igualment pero Get-InformeData les
# marca "ignorat" PER DEFECTE (nomes la primera vegada que es veu l'informe;
# vegeu _ConclusioIgnorarPerDefecte), perque son clausules gairebe identiques
# entre informes diferents (aporten poca informacio diferenciada per
# activitat).
$Script:ConclusioStartPhrases = @(
    [pscustomobject]@{ Font = 'vist_anterior'; Phrase = "Vist l'anterior" },
    [pscustomobject]@{ Font = 'risc';          Phrase = 'Tenint en consideració el risc' },
    [pscustomobject]@{ Font = 'mns';           Phrase = "S'informa favorablement" },
    [pscustomobject]@{ Font = 'act_extr';      Phrase = "El titular és responsable d'executar" },
    [pscustomobject]@{ Font = 'act_extr';      Phrase = "L'organitzador és responsable d'executar" }
)

# Conclusio: des del primer paragraf que conte una de $Script:ConclusioStartPhrases
# fins (exclos) el que marca el tancament de l'informe (signatura). Uneix els
# paragrafs amb un espai. Retorna un objecte { Text; Font }: Text es el text
# ORIGINAL (no normalitzat), '' si no es troba cap frase d'inici coneguda;
# Font indica quina frase ha disparat la deteccio (vegeu $Script:ConclusioStartPhrases).
function _ExtractConclusio($lines) {
    $starts = $Script:ConclusioStartPhrases | ForEach-Object {
        [pscustomobject]@{ Font = $_.Font; Norm = (_ConclNorm $_.Phrase) }
    }
    $endPhrases  = @(
        (_ConclNorm 'Ho poso al seu coneixement'),
        (_ConclNorm 'Cornella de Llobregat,'),
        (_ConclNorm "S'informa als efectes oportuns,"),
        (_ConclNorm 'A Cornella de Llobregat, en la data')
    )
    $start = -1
    $font = ''
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $nl = _ConclNorm $lines[$i]
        foreach ($sp in $starts) {
            if ($nl.Contains($sp.Norm)) { $start = $i; $font = $sp.Font; break }
        }
        if ($start -ge 0) { break }
    }
    if ($start -lt 0) { return [pscustomobject]@{ Text = ''; Font = '' } }
    $parts = New-Object System.Collections.ArrayList
    for ($i = $start; $i -lt $lines.Count; $i++) {
        $ln = [string]$lines[$i]
        $nl = _ConclNorm $ln
        $isEnd = $false
        foreach ($ep in $endPhrases) { if ($nl.Contains($ep)) { $isEnd = $true; break } }
        if ($isEnd) { break }
        if (-not [string]::IsNullOrWhiteSpace($ln)) { [void]$parts.Add($ln.Trim()) }
    }
    return [pscustomobject]@{ Text = ($parts -join ' '); Font = $font }
}

# Motiu de revisio associat a la conclusio detectada per _ExtractConclusio:
# 'sense conclusio' si no se n'ha trobat cap, '' si n'hi ha. Funcio PURA per
# poder-la testejar sense dependre de la lectura del document.
function _ConclusioMotiu($conclInfo) {
    if ([string]::IsNullOrWhiteSpace($conclInfo.Text)) { return 'sense conclusio' }
    return ''
}

# Cert si la conclusio detectada ve d'una familia de frases poc diferenciades
# entre informes ('mns', 'act_extr': gairebe la mateixa clausula sempre).
# Get-InformeData fa servir aixo per marcar l'informe "ignorat" PER DEFECTE
# nomes la primera vegada que es veu (a Invoke-InformesDbScan, si l'informe ja
# existia a un escaneig anterior es conserva l'"ignorat" que hi hagi marcat
# l'usuari, encara que l'informe s'hagi hagut de reprocessar).
function _ConclusioIgnorarPerDefecte($conclInfo) {
    return ($conclInfo.Font -eq 'mns' -or $conclInfo.Font -eq 'act_extr')
}

# Opcions valides de "conclusio breu" (l'estat en que queda l'activitat
# despres d'aquell informe). 'Altres' i 'Revisar' son manuals: la deteccio
# automatica (_ConclusioBreu) MAI les retorna directament com a resultat
# "trobat" -- nomes 'Revisar' com a valor per defecte quan no hi ha prou
# senyal. L'usuari les pot triar a ma des de l'editor si cal.
$Script:ConclusioBreuOpcions = @(
    'Requeriment',
    'FI Requeriment',
    'Precinte / Cessament',
    'FI Precinte / Cessament',
    'Favorable',
    'Ampliació termini',
    'Sense efecte',
    'Altres',
    'Revisar'
)

# Classifica el text de la CONCLUSIO (ja extreta per _ExtractConclusio) en un
# dels $Script:ConclusioBreuOpcions, mirant les frases reals amb que Sergi
# tanca cada tipus de tramit (vist a la carpeta real d'informes). 'Revisar' es
# el resultat per defecte quan no hi ha conclusio o no es reconeix cap frase
# (inclou "desfavorable", deliberadament: no es vol confondre amb "Favorable").
# Funcio PURA (nomes text), testejable en headless.
function _ConclusioBreu($text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return 'Revisar' }
    $n = _ConclNorm $text

    # Seguiment d'un requeriment: encara pendent. Qualsevol negacio de "es pot
    # donar per ..." (finalitzat / tancat / tancada la denuncia...) vol dir que
    # l'expedient NO es pot tancar: es un requeriment, no un FI. Ha d'anar ABANS
    # dels FI de sota (que fan servir la mateixa expressio sense el "no").
    if ($n -match "no s.?han esmenat" -or $n.Contains('no es pot donar')) { return 'Requeriment' }
    # Seguiment d'un requeriment: resolt (inclou denuncies tancades: mateix "final positiu").
    if ($n -match 'es pot donar.{0,12}finalitzat') { return 'FI Requeriment' }
    if ($n.Contains('es pot donar per tancada la denuncia')) { return 'FI Requeriment' }
    # Aixecament d'un precinte/suspensio.
    if ($n -match 'es (pot|valora) (aixecar|desprecintar)' -or $n.Contains('pertinent desprecintar')) { return 'FI Precinte / Cessament' }
    # Comunicacio anul·lada.
    if ($n.Contains('deixa sense efecte')) { return 'Sense efecte' }
    # Risc greu/imminent: es precinta o es proposa el cessament.
    if ($n.Contains('pertinent precintar') -or $n.Contains('tenint en consideracio el risc') -or $n -match 'ordeni el cessament') { return 'Precinte / Cessament' }
    # Desfavorable: deliberadament NO es classifica com a Favorable; cau a Revisar.
    if ($n.Contains('desfavorablement') -or $n.Contains('desfavorable')) { return 'Revisar' }
    if ($n.Contains('favorablement') -or $n.Contains('favorable')) { return 'Favorable' }
    if ($n.Contains('ampliar el termini')) { return 'Ampliació termini' }
    # Clausules estandard d'un requeriment NOU (encara sense "Vist l'anterior").
    if ($n.Contains('recepcio del requeriment') -or $n.Contains('esmenar les deficiencies') -or
        $n.Contains('mancances formals') -or $n.Contains('termini maxim de') -or
        $n.Contains('podran adoptar les mesures') -or
        $n.Contains('procediment desmena') -or $n.Contains('esmenar els defectes') -or
        $n -match 'cas contrari.{0,60}(cessament|precinte)' -or $n -match 'determini el (cessament|precinte)') {
        return 'Requeriment'
    }
    return 'Revisar'
}

# Estat actual d'una ACTIVITAT: la conclusio_breu del seu informe mes RECENT
# (per data) entre els que NO estan ignorats (un informe ignorat -p.ex. una
# clausula MNS/act_extr poc fiable, o marcat a ma- no ha de decidir l'estat
# de l'activitat). Espera $informesOrdenats ja ordenats per data ASCENDENT
# (com fa Invoke-InformesDbScan); pren el darrer que compleixi. '' si no n'hi
# ha cap (activitat sense cap informe fiable). Funcio PURA, testejable.
# L'informe que DECIDEIX l'estat de l'activitat: l'ULTIM dels fiables (els
# informes venen ordenats per data). $null si no n'hi ha cap.
# Es la font unica: _EstatActualActivitat i qui vulgui saber-ne la DATA
# (p. ex. "Comprovar Excel") criden aquesta mateixa funcio, aixi no poden
# discrepar mai sobre quin informe manda.
function _InformeQueDeterminaEstat($informesOrdenats) {
    if ($null -eq $informesOrdenats) { return $null }
    $fiables = @($informesOrdenats | Where-Object { -not [bool]$_.ignorat })
    if ($fiables.Count -eq 0) { return $null }
    return $fiables[-1]
}

function _EstatActualActivitat($informesOrdenats) {
    $inf = _InformeQueDeterminaEstat $informesOrdenats
    if ($null -eq $inf) { return '' }
    return [string]$inf.conclusio_breu
}

# ----------------------------------------------------------------------------
# EDICIONS A MA ("Editar base") -- octubre 2026, peticio de l'usuari: el que
# s'ha corregit a ma PREVAL sobre el que surti d'"Actualitzar base", i l'estat
# d'aquelles activitats surt en VERMELL a l'editor perque se'n sigui conscient.
#
# Es marca a l'INFORME (no a l'activitat): editat_a_ma = $true, i es guarda el
# valor automatic que hi havia (auto_conclusio_breu / auto_ignorat) per poder
# desfer-ho. Abans de la marca, l'escaneig ja conservava la conclusio breu i
# l'ignorat de TOTS els informes reprocessats -editats o no-, o sigui que un
# informe que es tornava a escriure mai no actualitzava la seva conclusio breu.
# ----------------------------------------------------------------------------
function _PropInf($o, [string]$nom) {
    if ($null -eq $o -or $null -eq $o.PSObject.Properties[$nom]) { return $null }
    return $o.$nom
}

# Marca un informe com a editat a ma, guardant ABANS el valor automatic. Si ja
# ho estava, no toca res (el valor automatic que es guarda es el primer).
function _MarcaEditatAMa($inf) {
    if ([bool](_PropInf $inf 'editat_a_ma')) { return }
    Add-Member -InputObject $inf -NotePropertyName auto_conclusio_breu -NotePropertyValue ([string](_PropInf $inf 'conclusio_breu')) -Force
    Add-Member -InputObject $inf -NotePropertyName auto_ignorat -NotePropertyValue ([bool](_PropInf $inf 'ignorat')) -Force
    Add-Member -InputObject $inf -NotePropertyName editat_a_ma -NotePropertyValue $true -Force
}

# Torna un informe al que diu l'automatic. Retorna $true si ho ha fet.
function _DesfesEditatAMa($inf) {
    if (-not [bool](_PropInf $inf 'editat_a_ma')) { return $false }
    Add-Member -InputObject $inf -NotePropertyName conclusio_breu -NotePropertyValue ([string](_PropInf $inf 'auto_conclusio_breu')) -Force
    Add-Member -InputObject $inf -NotePropertyName ignorat -NotePropertyValue ([bool](_PropInf $inf 'auto_ignorat')) -Force
    Add-Member -InputObject $inf -NotePropertyName editat_a_ma -NotePropertyValue $false -Force
    [void]$inf.PSObject.Properties.Remove('auto_conclusio_breu')
    [void]$inf.PSObject.Properties.Remove('auto_ignorat')
    return $true
}

# Una base d'ABANS de la marca: un informe la conclusio breu del qual no es la
# que en surt del text NOMES pot ser una correccio a ma (l'automatic sempre la
# treu del text). Se li posa la marca, amb el valor automatic que en surt.
# L'ignorat, aqui, no es pot deduir (el per defecte depen de coses que la base
# no guarda); el dedueix l'escaneig, que si que les te. Un informe que ja porta
# la marca (encara que sigui $false) no es toca.
function _InferEditatAMa($inf) {
    if ($null -ne $inf.PSObject.Properties['editat_a_ma']) { return }
    $auto = _ConclusioBreu ([string](_PropInf $inf 'conclusio'))
    $actual = [string](_PropInf $inf 'conclusio_breu')
    if ($actual -ne '' -and $actual -ne $auto) {
        Add-Member -InputObject $inf -NotePropertyName auto_conclusio_breu -NotePropertyValue $auto -Force
        Add-Member -InputObject $inf -NotePropertyName auto_ignorat -NotePropertyValue ([bool](_PropInf $inf 'ignorat')) -Force
        Add-Member -InputObject $inf -NotePropertyName editat_a_ma -NotePropertyValue $true -Force
    } else {
        Add-Member -InputObject $inf -NotePropertyName editat_a_ma -NotePropertyValue $false -Force
    }
}

# Alguna conclusio/ignorat d'aquesta activitat s'ha corregit a ma?
function _ActivitatEditadaAMa($act) {
    foreach ($inf in @(_PropInf $act 'informes')) { if ($null -ne $inf -and [bool](_PropInf $inf 'editat_a_ma')) { return $true } }
    return $false
}


# L'ID GIA com a NUMERO per ordenar: '9' abans que '10' (com a text, '10' anava
# abans que '9' i la llista sortia 10, 1000, 1019, 103...). Sense GIA o no
# numeric, al final.
function _GiaNumeric($gia) {
    $n = 0L
    if ([long]::TryParse(([string]$gia).Trim(), [ref]$n)) { return $n }
    return [long]::MaxValue
}

# L'ordre de les files de l'editor. PURA.
#   Sense columna triada: per activitat (ID GIA NUMERIC; les que no en tenen, al
#   final i per carpeta) i dins de cada una, per data.
#   Amb una columna triada (clic a la capcalera): AQUELLA COLUMNA MANA, i l'ID
#   GIA i la data nomes desempaten. Abans l'agrupament per activitat era sempre
#   la clau primaria i la columna nomes ordenava DINS de cada activitat, o sigui
#   que clicar 'Estat activitat' no ordenava res que es veies.
# $colExpr: hashtable index de columna -> scriptblock sobre la fila ($_).
function _OrdenaFilesBase($rows, $colExpr, [int]$sortCol, [bool]$asc) {
    $crit = @()
    if ($sortCol -ge 0 -and $null -ne $colExpr -and $colExpr.ContainsKey($sortCol)) {
        $crit += @{ Expression = $colExpr[$sortCol]; Descending = (-not $asc) }
    }
    $crit += @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.Gia)) { 1 } else { 0 } } }
    $crit += @{ Expression = { _GiaNumeric $_.Gia } }
    $crit += @{ Expression = { [string]$_.Carpeta } }
    $crit += @{ Expression = { [string]$_.Data } }
    return @(@($rows) | Sort-Object -Property $crit)
}

# La data d'un informe en format de carrer: 'yyyy-MM-dd' -> 'dd/MM/yyyy'.
# Buit si no hi ha data o no te aquest format (mai peta). Funcio PURA.
function _DataInformeDdMmAaaa($data) {
    $s = [string]$data
    if ($s -match '^(\d{4})-(\d{2})-(\d{2})') {
        return ('{0}/{1}/{2}' -f $Matches[3], $Matches[2], $Matches[1])
    }
    return ''
}

# ID GIA a partir dels noms de les carpetes pare (p.ex. la carpeta de l'activitat
# "2025-1-2563 GIA 361 - ... KRICHI ..."). Retorna '' si cap carpeta no en porta.
# Parteix la ruta manualment per '\' o '/' (aixi funciona igual a Windows i a
# Linux, i es pot provar en headless).
function _GiaFromFolderName($path) {
    $segs = ([string]$path) -split '[\\/]'
    # L'ultim segment es el nom del fitxer; recorrem les carpetes de dins enfora.
    for ($i = $segs.Count - 2; $i -ge 0; $i--) {
        if ($segs[$i] -match 'GIA\s*(\d+)') { return $Matches[1] }
    }
    return ''
}

# Nom de la carpeta (activitat) on viu l'informe: la carpeta pare immediata.
function _CarpetaActivitat($path) {
    $segs = ([string]$path) -split '[\\/]' | Where-Object { $_ -ne '' }
    if ($segs.Count -ge 2) { return $segs[$segs.Count - 2] }
    return ''
}

# Normalitza un numero d'expedient per comparar-lo (l'informe fa servir "/", la
# carpeta "-", i l'Excel pot portar zeros al davant): parteix en grups i treu els
# zeros inicials de cada grup numeric. "2025/1/2563" i "2025/01/2563" -> "2025-1-2563".
function _NormalitzaExpedient($s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    $groups = ([string]$s).Trim() -split '[^\dA-Za-z]+' | Where-Object { $_ -ne '' }
    $norm = $groups | ForEach-Object {
        if ($_ -match '^\d+$') { [string][int]$_ } else { $_.ToUpper() }
    }
    return ($norm -join '-')
}

# Construeix un mapa expedient_normalitzat -> ID GIA a partir de la cache de
# l'Excel d'activitats ($cache.ById[id] = @{ EXP_NUM; TITULAR; ... }).
function Build-ExpedientToGiaMap($cache) {
    $map = @{}
    if ($null -eq $cache -or $null -eq $cache.ById) { return $map }
    foreach ($kv in $cache.ById.GetEnumerator()) {
        $exp = _NormalitzaExpedient $kv.Value.EXP_NUM
        if ($exp -ne '' -and -not $map.ContainsKey($exp)) { $map[$exp] = [string]$kv.Key }
    }
    return $map
}


# ----------------------------------------------------------------------------
# Editor de la base d'informes (finestra amb taula)
# ----------------------------------------------------------------------------
# Estil d'una fila segons si l'informe esta ignorat: gris + tatxat si ho esta.
function _StyleInformeRow($gridRow, [bool]$ignorat, $fontNormal, $fontStrike) {
    if ($ignorat) {
        $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
        $gridRow.DefaultCellStyle.Font      = $fontStrike
    } else {
        $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
        $gridRow.DefaultCellStyle.Font      = $fontNormal
    }
}

# Obre una finestra amb la base d'informes en forma de taula: es pot veure cada
# informe (data, GIA, titular, carpeta, conclusio), OBRIR-lo (boto), marcar-lo
# com a IGNORAT (casella; reversible) i corregir-ne la CONCLUSIO BREU
# (desplegable editable). La columna "Estat activitat" es NOMES LECTURA: es
# deriva de la conclusio breu del darrer informe no ignorat de l'activitat
# (per data) i es recalcula sempre que canvia "ignorar" o "conclusio breu" de
# qualsevol dels seus informes. Tots els canvis es desen al JSON.
function Invoke-InformesDbEdit {
    $outPath = Join-Path $LocalActivitatsDir 'informes-db.json'
    if (-not (Test-Path -LiteralPath $outPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Encara no hi ha cap base d'informes.`n`nExecuta primer 'Actualitzar base d'informes'.",
            'Editar base d''informes', 'OK', 'Information') | Out-Null
        return
    }
    $db = $null
    try {
        $db = Read-JsonFile $outPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir la base:`n$($_.Exception.Message)", 'Editar base d''informes', 'OK', 'Error') | Out-Null
        return
    }

    # Aplanar en files, normalitzant cada informe (assegurar ignorat/ruta/motiu/
    # conclusio_breu). Cada fila guarda una REFERENCIA a l'objecte informe del
    # JSON ($inf) i a l'activitat ($act), aixi editar "ignorar" o "conclusio
    # breu" modifica directament la base que despres desem, i podem recalcular
    # l'"estat actual" de l'activitat (sempre DERIVAT: el recomputem en carregar
    # i cada cop que canvia algun dels seus informes, mai es desa "a cegues").
    $allRows = New-Object System.Collections.ArrayList
    if ($null -ne $db.activitats) {
        foreach ($act in $db.activitats) {
            if ($null -eq $act.informes) { continue }
            foreach ($inf in $act.informes) {
                if ($null -eq $inf.PSObject.Properties['ignorat'])        { Add-Member -InputObject $inf -NotePropertyName ignorat -NotePropertyValue $false -Force }
                if ($null -eq $inf.PSObject.Properties['ruta'])           { Add-Member -InputObject $inf -NotePropertyName ruta -NotePropertyValue '' -Force }
                if ($null -eq $inf.PSObject.Properties['motiu'])          { Add-Member -InputObject $inf -NotePropertyName motiu -NotePropertyValue '' -Force }
                if ($null -eq $inf.PSObject.Properties['conclusio_breu']) { Add-Member -InputObject $inf -NotePropertyName conclusio_breu -NotePropertyValue (_ConclusioBreu ([string]$inf.conclusio)) -Force }
                _InferEditatAMa $inf
            }
            if ($null -eq $act.PSObject.Properties['estat_actual']) { Add-Member -InputObject $act -NotePropertyName estat_actual -NotePropertyValue '' -Force }
            $act.estat_actual = _EstatActualActivitat $act.informes
            foreach ($inf in $act.informes) {
                [void]$allRows.Add([pscustomobject]@{
                    Obj           = $inf
                    Act           = $act
                    Data          = [string]$inf.data
                    Gia           = [string]$act.id_gia
                    Titular       = [string]$act.titular
                    Carpeta       = [string]$act.carpeta
                    Conclusio     = [string]$inf.conclusio
                    ConclusioBreu = [string]$inf.conclusio_breu
                    EstatActual   = [string]$act.estat_actual
                    Motiu         = [string]$inf.motiu
                    Ruta          = [string]$inf.ruta
                })
            }
        }
    }

    # Per activitat (ID GIA NUMERIC; les que no en tenen, al final i per carpeta)
    # i per data: vegeu _OrdenaFilesBase.
    $allRows = _OrdenaFilesBase $allRows @{} -1 $true

    # Estat compartit amb els gestors d'esdeveniments (hashtable per referencia).
    # 'Loading' evita que el gestor de la casella reaccioni mentre s'omple la
    # graella (Rows.Add pot disparar CellValueChanged abans d'assignar el Tag).
    # SortColIdx/SortAsc: la columna que l'usuari ha clicat (-1 = cap). Mana
    # ella, i l'ID GIA i la data nomes desempaten (_OrdenaFilesBase).
    $state = @{ Dirty = $false; Db = $db; Path = $outPath; Loading = $false; SortColIdx = -1; SortAsc = $true }

    $form = _NewForm
    $form.Text = "Editar base d'informes"
    $form.Size = New-Object System.Drawing.Size(1000, 676)
    $form.MinimumSize = New-Object System.Drawing.Size(720, 476)

    # Graella.
    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid

    $cData = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cData.HeaderText = 'Data';      $cData.ReadOnly = $true; $cData.Width = 90
    $cGia  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cGia.HeaderText  = 'GIA';       $cGia.ReadOnly  = $true; $cGia.Width  = 60
    $cTit  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cTit.HeaderText  = 'Titular';   $cTit.ReadOnly  = $true; $cTit.Width  = 180
    $cCar  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cCar.HeaderText  = 'Carpeta';   $cCar.ReadOnly  = $true; $cCar.Width  = 190
    $cCon  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cCon.HeaderText  = 'Conclusio'; $cCon.ReadOnly  = $true; $cCon.Width  = 260
    $cBreu = New-Object System.Windows.Forms.DataGridViewComboBoxColumn; $cBreu.HeaderText = 'Conclusio breu'; $cBreu.Width = 150; $cBreu.FlatStyle = 'Flat'
    [void]$cBreu.Items.AddRange($Script:ConclusioBreuOpcions)
    $cEst  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cEst.HeaderText  = 'Estat activitat'; $cEst.ReadOnly = $true; $cEst.Width = 150
    $cMot  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cMot.HeaderText  = 'Motiu';     $cMot.ReadOnly  = $true; $cMot.Width  = 110
    $cObr  = New-Object System.Windows.Forms.DataGridViewButtonColumn;   $cObr.HeaderText  = '';          $cObr.Text = 'Obrir'; $cObr.UseColumnTextForButtonValue = $true; $cObr.Width = 64
    $cIgn  = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $cIgn.HeaderText  = 'Ignorar';   $cIgn.Width = 60
    [void]$grid.Columns.Add($cData)
    [void]$grid.Columns.Add($cGia)
    [void]$grid.Columns.Add($cTit)
    [void]$grid.Columns.Add($cCar)
    [void]$grid.Columns.Add($cCon)
    [void]$grid.Columns.Add($cBreu)
    [void]$grid.Columns.Add($cEst)
    [void]$grid.Columns.Add($cMot)
    [void]$grid.Columns.Add($cObr)
    [void]$grid.Columns.Add($cIgn)
    $idxConclBreu = 5
    $idxEstat     = 6
    $idxObrir     = 8
    $idxIgnorar   = 9

    # Ordenacio PROGRAMATICA: capturem el clic a la capcalera nosaltres mateixos
    # (mes avall): el DataGridView ordenaria l'ID GIA com a text i no desempataria
    # per activitat. Desactivem l'ordenacio automatica de totes les columnes.
    foreach ($col in $grid.Columns) {
        if ($col.Index -ne $idxObrir) { $col.SortMode = 'Programmatic' }
    }

    $fontNormal = $grid.Font
    $fontStrike = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Strikeout)
    $fontBold   = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Bold)
    $colEditat  = [System.Drawing.Color]::FromArgb(192, 0, 0)

    # L'ESTAT EN VERMELL quan alguna cosa de l'activitat s'ha corregit a ma: es
    # el que preval sobre "Actualitzar base", i l'usuari ho vol veure.
    $pintaEstat = {
        param($gr, $row)
        $c = $gr.Cells[$idxEstat]
        if (_ActivitatEditadaAMa $row.Act) {
            $c.Style.ForeColor = $colEditat
            $c.Style.Font = $fontBold
            $c.ToolTipText = "Corregida a m" + [char]0x00E0 + " ('Conclusi" + [char]0x00F3 + " breu' o 'Ignorar'): 'Actualitzar base' no la canvia. Per tornar al que diu l'informe, selecciona la fila i 'Desfer canvi a m" + [char]0x00E0 + "'."
        } else {
            $c.Style.ForeColor = [System.Drawing.Color]::Empty
            $c.Style.Font = $null
            $c.ToolTipText = ''
        }
    }.GetNewClosure()

    # Torna a escriure l'estat de TOTES les files d'una activitat (es el que
    # canvia quan canvia qualsevol dels seus informes).
    $refrescaActivitat = {
        param($act)
        $nouEstat = _EstatActualActivitat $act.informes
        $act.estat_actual = $nouEstat
        foreach ($gr2 in $grid.Rows) {
            $row2 = $gr2.Tag
            if ($null -ne $row2 -and $row2.Act -eq $act) {
                $row2.EstatActual = $nouEstat
                $gr2.Cells[$idxEstat].Value = $nouEstat
                & $pintaEstat $gr2 $row2
            }
        }
        foreach ($row3 in $allRows) {
            if ($row3.Act -eq $act) { $row3.EstatActual = $nouEstat }
        }
    }.GetNewClosure()

    # Expressions d'ordenacio per index de columna (sobre la fila $_). El boto
    # "Obrir" (8) no s'ordena.
    $colExpr = @{
        0 = { [string]$_.Data }
        1 = { _GiaNumeric $_.Gia }
        2 = { [string]$_.Titular }
        3 = { [string]$_.Carpeta }
        4 = { [string]$_.Conclusio }
        5 = { [string]$_.ConclusioBreu }
        6 = { [string]$_.EstatActual }
        7 = { [string]$_.Motiu }
        9 = { [bool]$_.Obj.ignorat }
    }

    # ---- Barra superior: cerca global (fila 1) + filtres per columna (fila 2) ----
    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = 'Top'; $topPanel.Height = 74
    # La cerca torna a omplir la graella. ($fill es defineix mes avall; ja
    # existeix quan l'usuari hi escriu.)
    $txtCerca = _AddSearchBox $topPanel 10 10 300 'Cerca:' { & $fill }

    # Etiqueta a la fila 2 dels filtres.
    $mkLbl = {
        param($text, $x)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.AutoSize = $true
        $l.Location = New-Object System.Drawing.Point($x, 46)
        $topPanel.Controls.Add($l)
    }

    $estatVals = @($allRows | ForEach-Object { $_.EstatActual } | Where-Object { $_ -ne '' } | Sort-Object -Unique)
    $motiuVals = @($allRows | ForEach-Object { $_.Motiu }       | Where-Object { $_ -ne '' } | Sort-Object -Unique)
    # Filtres de SELECCIO MULTIPLE (cap marcat = totes les files passen). L'accio
    # de canvi crida $fill (definit mes avall; ja existeix quan l'usuari hi toca).
    & $mkLbl 'Conclusió breu:' 10
    $mfBreu  = _MakeMultiFilter $topPanel 110 43 130 '(Totes)' $Script:ConclusioBreuOpcions { & $fill }
    & $mkLbl 'Estat:' 250
    $mfEstat = _MakeMultiFilter $topPanel 294 43 120 '(Tots)' $estatVals { & $fill }
    & $mkLbl 'Motiu:' 420
    $mfMotiu = _MakeMultiFilter $topPanel 464 43 110 '(Tots)' $motiuVals { & $fill }
    & $mkLbl 'Ignorats:' 582
    $mfIgn   = _MakeMultiFilter $topPanel 640 43 90  'Tots'    @('Actius', 'Ignorats') { & $fill }

    # (Re)omple la graella: aplica la cerca global + els filtres per columna i,
    # per acabar, ordena mantenint SEMPRE l'agrupament per activitat.
    $fill = {
        $state.Loading = $true
        $grid.Rows.Clear()
        $n = ([string]$txtCerca.Text).Trim().ToLower()
        $selBreu  = & $mfBreu.GetSelected
        $selEstat = & $mfEstat.GetSelected
        $selMotiu = & $mfMotiu.GetSelected
        $selIgn   = & $mfIgn.GetSelected

        $rows = foreach ($row in $allRows) {
            $hay = $row.Data + ' ' + $row.Gia + ' ' + $row.Titular + ' ' + $row.Carpeta + ' ' + $row.Conclusio + ' ' + $row.ConclusioBreu + ' ' + $row.EstatActual + ' ' + $row.Motiu
            if (-not (_TextMatches $hay $n)) { continue }
            # Cada filtre: cap opcio marcada = passa tot; si n'hi ha, el valor de
            # la fila ha de ser entre les marcades (unio / OR dins del filtre).
            if ($selBreu.Count  -gt 0 -and $selBreu  -notcontains $row.ConclusioBreu) { continue }
            if ($selEstat.Count -gt 0 -and $selEstat -notcontains $row.EstatActual)   { continue }
            if ($selMotiu.Count -gt 0 -and $selMotiu -notcontains $row.Motiu)         { continue }
            if ($selIgn.Count -gt 0) {
                $ignLabel = if ([bool]$row.Obj.ignorat) { 'Ignorats' } else { 'Actius' }
                if ($selIgn -notcontains $ignLabel) { continue }
            }
            $row
        }
        $rows = @($rows)

        $rows = _OrdenaFilesBase $rows $colExpr ([int]$state.SortColIdx) ([bool]$state.SortAsc)

        foreach ($row in $rows) {
            $ign = [bool]$row.Obj.ignorat
            $idx = $grid.Rows.Add(@($row.Data, $row.Gia, $row.Titular, $row.Carpeta, $row.Conclusio, $row.ConclusioBreu, $row.EstatActual, $row.Motiu, 'Obrir', $ign))
            $gr = $grid.Rows[$idx]
            $gr.Tag = $row
            $gr.Cells[4].ToolTipText = $row.Conclusio
            _StyleInformeRow $gr $ign $fontNormal $fontStrike
            & $pintaEstat $gr $row
        }
        $state.Loading = $false
    }.GetNewClosure()

    # Clic a la capcalera: tria la columna que mana i alterna asc/desc. La
    # columna "Obrir" (boto) no s'ordena.
    _EnableHeaderSort $grid $state 0 @($idxObrir) { & $fill }

    # Desa la base (retorna $true si va be).
    $doSave = {
        try {
            Write-JsonFile $state.Path $state.Db 8
            $state.Dirty = $false
            return $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n$($_.Exception.Message)", 'Editar base d''informes', 'OK', 'Error') | Out-Null
            return $false
        }
    }.GetNewClosure()

    # "Desfer canvi a ma": les files seleccionades tornen al que diu l'informe.
    # Definit ABANS del peu de botons: el boto en captura el valor en crear-se.
    $desfesAMa = {
        $files = @($grid.SelectedRows)
        if ($files.Count -eq 0 -and $null -ne $grid.CurrentRow) { $files = @($grid.CurrentRow) }
        $n = 0
        $state.Loading = $true
        try {
            foreach ($gr in $files) {
                $row = $gr.Tag
                if ($null -eq $row -or -not (_DesfesEditatAMa $row.Obj)) { continue }
                $n++
                $row.ConclusioBreu = [string]$row.Obj.conclusio_breu
                $gr.Cells[$idxConclBreu].Value = $row.ConclusioBreu
                $gr.Cells[$idxIgnorar].Value = [bool]$row.Obj.ignorat
                _StyleInformeRow $gr ([bool]$row.Obj.ignorat) $fontNormal $fontStrike
                & $refrescaActivitat $row.Act
            }
        } finally { $state.Loading = $false }
        if ($n -gt 0) { $state.Dirty = $true }
        else {
            [System.Windows.Forms.MessageBox]::Show("La fila seleccionada no t" + [char]0x00E9 + " cap canvi fet a m" + [char]0x00E0 + ".", 'Editar base d''informes', 'OK', 'Information') | Out-Null
        }
    }.GetNewClosure()

    # Barra inferior: botons.
    $botPanel = New-Object System.Windows.Forms.Panel
    $botPanel.Dock = 'Bottom'; $botPanel.Height = 48
    # Exportar a CSV els llistats d'activitats en Estat Requeriment i Precinte /
    # Cessament (usa l'estat en memoria, que ja reflecteix els canvis no desats).
    [void](_AddPeuBotons $form @(
        @{ Nom = 'Enrere'; Text = (_TxtEnrere); Clic = { $form.Close() }.GetNewClosure() },
        @{ Nom = 'Export'; Text = 'Exportar llistats (CSV)'; Clic = { Export-EstatsActivitats $state.Db }.GetNewClosure() },
        @{ Nom = 'Desfer'; Text = ('Desfer canvi a m' + [char]0x00E0); Clic = { & $desfesAMa }.GetNewClosure() }) @(
        @{ Nom = 'Desar'; Text = 'Desar'; Estil = 'primari'; Clic = {
            if (& $doSave) {
                [System.Windows.Forms.MessageBox]::Show('Canvis desats.', 'Editar base d''informes', 'OK', 'Information') | Out-Null
            }
        }.GetNewClosure() }) 8 $botPanel)

    # Obrir l'informe en clicar el boto "Obrir".
    $grid.add_CellContentClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -eq $idxObrir) {
            $row = $s.Rows[$e.RowIndex].Tag
            $ruta = [string]$row.Ruta
            if ([string]::IsNullOrWhiteSpace($ruta) -or -not (Test-Path -LiteralPath $ruta)) {
                [System.Windows.Forms.MessageBox]::Show("No s'ha trobat el fitxer:`n$ruta", 'Obrir informe', 'OK', 'Warning') | Out-Null
                return
            }
            try { Start-Process -FilePath $ruta | Out-Null } catch {
                [System.Windows.Forms.MessageBox]::Show("No s'ha pogut obrir:`n$($_.Exception.Message)", 'Obrir informe', 'OK', 'Error') | Out-Null
            }
        }
    }.GetNewClosure())

    # Forcem que la casella "Ignorar" es confirmi de seguida (no en sortir de la
    # cel-la), perque CellValueChanged salti al moment del clic.
    $grid.add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty) {
            $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    }.GetNewClosure())
    $grid.add_CellValueChanged({
        param($s, $e)
        if ($state.Loading) { return }
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -ne $idxIgnorar -and $e.ColumnIndex -ne $idxConclBreu) { return }
        $gr = $s.Rows[$e.RowIndex]
        $row = $gr.Tag
        if ($null -eq $row) { return }

        # Un canvi a ma: es marca ABANS de canviar el valor, perque s'hi guarda
        # el que deia l'automatic (per poder-ho desfer).
        _MarcaEditatAMa $row.Obj
        if ($e.ColumnIndex -eq $idxIgnorar) {
            $val = [bool]$gr.Cells[$idxIgnorar].Value
            $row.Obj.ignorat = $val
            _StyleInformeRow $gr $val $fontNormal $fontStrike
        } else {
            $val = [string]$gr.Cells[$idxConclBreu].Value
            if ([string]::IsNullOrWhiteSpace($val)) { $val = 'Revisar' }
            $row.Obj.conclusio_breu = $val
            $row.ConclusioBreu = $val
        }
        $state.Dirty = $true

        # L'"ignorar" i la "conclusio breu" de qualsevol informe poden canviar
        # l'estat de l'activitat sencera: el recalculem i el propaguem a totes
        # les files (visibles i filtrades) d'aquesta mateixa activitat.
        & $refrescaActivitat $row.Act
    }.GetNewClosure())


    # Si es tanca amb canvis sense desar, oferim desar-los.
    $form.add_FormClosing({
        param($s, $e)
        if ($state.Dirty) {
            $r = [System.Windows.Forms.MessageBox]::Show('Hi ha canvis sense desar. Vols desar-los?', 'Editar base d''informes', 'YesNoCancel', 'Warning')
            if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
                if (-not (& $doSave)) { $e.Cancel = $true }
            } elseif ($r -eq [System.Windows.Forms.DialogResult]::Cancel) {
                $e.Cancel = $true
            }
        }
    }.GetNewClosure())

    # Ordre d'afegit: primer el Fill (queda al centre), despres Top i Bottom.
    # La banda de marca s'afegeix l'ultima perque quedi a dalt de tot.
    $form.Controls.Add($grid)
    $form.Controls.Add($topPanel)
    $form.Controls.Add($botPanel)
    [void](_AddBrandHeader $form "Editar base d'informes" 'Base de dades local de deficiencies i conclusions dels informes' 56)

    & $fill
    [void]$form.ShowDialog()
}

# ============================================================================
# Exportar llistats (CSV) — botó "Exportar llistats" de l'editor
# ============================================================================
# Genera un CSV amb les activitats en Estat "Requeriment" i "Precinte /
# Cessament" (una fila per informe). Creua amb l'Excel per obtenir l'adreça
# (si no hi ha Excel, l'adreça queda buida). S'obre a Excel en acabar.
function Export-EstatsActivitats($db) {
    if ($null -eq $db -or $null -eq $db.activitats) {
        [System.Windows.Forms.MessageBox]::Show("La base d'informes és buida.", 'Exportar llistats', 'OK', 'Information') | Out-Null
        return
    }
    $estats = @('Requeriment', 'Precinte / Cessament')
    $acts = @($db.activitats | Where-Object { $estats -contains [string]$_.estat_actual })
    if ($acts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No hi ha cap activitat en Estat 'Requeriment' ni 'Precinte / Cessament'.", 'Exportar llistats', 'OK', 'Information') | Out-Null
        return
    }

    # Adreça per GIA des de l'Excel (opcional; avís suau si no es pot llegir).
    $adrByGia = @{}
    try {
        $xls = Find-LatestActivitatsExcel
        if ($null -ne $xls) {
            $cache = Initialize-ActivitatsCache $xls.File
            if ($null -ne $cache -and $null -ne $cache.ById) {
                foreach ($kv in $cache.ById.GetEnumerator()) { $adrByGia[[string]$kv.Key] = [string]$kv.Value.ADRECA }
            }
        }
    } catch { }

    $rows = New-Object System.Collections.ArrayList
    foreach ($act in $acts) {
        $gia = [string]$act.id_gia
        $adr = if ($gia -and $adrByGia.ContainsKey($gia)) { $adrByGia[$gia] } else { '' }
        foreach ($inf in @($act.informes)) {
            if ($null -eq $inf) { continue }
            [void]$rows.Add([pscustomobject]@{
                'Estat activitat'  = [string]$act.estat_actual
                'GIA'              = $gia
                'Titular'          = [string]$act.titular
                'Adreça'           = $adr
                'Expedient'        = [string]$act.expedient
                'Data informe'     = [string]$inf.data
                'Conclusió breu'   = [string]$inf.conclusio_breu
            })
        }
    }
    $sorted = @($rows | Sort-Object 'Estat activitat', 'GIA', 'Data informe')

    $dir  = _ResolveOutputDir
    $name = 'Estat activitats ' + (Get-Date).ToString('yyyy-MM-dd') + '.csv'
    $path = _GetUniqueOutputPath $dir $name
    try {
        $sorted | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut escriure el CSV:`n$($_.Exception.Message)", 'Exportar llistats', 'OK', 'Error') | Out-Null
        return
    }

    $nReq = @($acts | Where-Object { [string]$_.estat_actual -eq 'Requeriment' }).Count
    $nPre = @($acts | Where-Object { [string]$_.estat_actual -eq 'Precinte / Cessament' }).Count
    $r = [System.Windows.Forms.MessageBox]::Show(
        "CSV generat:`n$path`n`nRequeriment: $nReq activitats`nPrecinte / Cessament: $nPre activitats`n`nVols obrir-lo ara?",
        'Exportar llistats', 'YesNo', 'Information')
    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
        try { Start-Process -FilePath $path | Out-Null } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut obrir el CSV:`n$($_.Exception.Message)", 'Exportar llistats', 'OK', 'Warning') | Out-Null
        }
    }
}
