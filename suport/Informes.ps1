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
  Ho desa AGRUPAT PER ACTIVITAT a BASE DE DADES ACTIVITATS\informes-db.json
  (carpeta ignorada per git). Els informes que no es poden resoldre del tot
  van a un bloc "a_revisar".

  Es un modul del motor: es carrega (dot-source) des de GenerarInforme.ps1, aixi
  reutilitza les funcions de lectura de .docx sense Word (de Seguiment.ps1),
  _NormalizeText i l'acces a l'Excel d'activitats (Find-LatestActivitatsExcel /
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
        $nt = _NormalizeText $s
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
    $t = _NormalizeText $s
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

# ----------------------------------------------------------------------------
# Comprovar Excel: noms de "Camp Info" que, a l'Excel d'activitats, marquen que
# una activitat esta requerida per decret / precintada. La comprovacio es fa
# sobre el NOM del camp (normalitzat, sense accents/majuscules; el '?' es
# conserva) i, a mes, el VALOR ha de comencar per "SI".
$Script:ExcelPrecinteCampNoms = @('requerit per decret?', 'precinte?')

# PURA i testejable. $pairs = llista de @{ Nom; Valor } (els Camp Info d'una fila
# de l'Excel). Retorna $true si algun te el Nom entre els objectius I el Valor
# comenca per "SI". _NormalizeText es de GenerarInforme.ps1 (ja dot-sourcejat).
function _ExcelActivitatActualitzada($pairs) {
    if ($null -eq $pairs) { return $false }
    foreach ($p in $pairs) {
        $nom = _NormalizeText ([string]$p.Nom)
        if ($Script:ExcelPrecinteCampNoms -contains $nom) {
            $val = _NormalizeText ([string]$p.Valor)
            if ($val -match '^si\b') { return $true }
        }
    }
    return $false
}

# PURA i testejable. A partir de la fila de capcaleres (array 0-based de textos),
# localitza els parells de columnes "Camp Info N - Nom" / "Camp Info N - Valor"
# i retorna una llista de @{ NomCol; ValorCol } (indexs 1-based, com Excel).
function _FindCampInfoPairs($headers) {
    $h = @($headers)
    $pairs = @()
    for ($i = 0; $i -lt $h.Count; $i++) {
        $n = _NormalizeText $h[$i]
        if ($n -match '^camp info\s+(\d+)\s*-\s*nom$') {
            $num = $Matches[1]
            $target = _NormalizeText ("camp info $num - valor")
            $valorCol = 0
            for ($j = 0; $j -lt $h.Count; $j++) {
                if ((_NormalizeText $h[$j]) -eq $target) { $valorCol = $j + 1; break }
            }
            if ($valorCol -gt 0) { $pairs += @{ NomCol = ($i + 1); ValorCol = $valorCol } }
        }
    }
    return ,@($pairs)
}

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
function _EstatActualActivitat($informesOrdenats) {
    if ($null -eq $informesOrdenats) { return '' }
    $fiables = @($informesOrdenats | Where-Object { -not [bool]$_.ignorat })
    if ($fiables.Count -eq 0) { return '' }
    return [string]$fiables[-1].conclusio_breu
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

# Analitza UN informe. Retorna un PSCustomObject amb data, gia, expedient,
# conclusio, fitxer, ruta, carpeta i el motiu (si cal revisar-lo). $wordApp es
# opcional (nomes cal per als .doc antics; vegeu _ReadInformeParagraphs).
function Get-InformeData($file, $expToGia, $cache, $wordApp = $null) {
    $data = _ParseDataInformeFromName $file.Name
    $lines = @()
    try { $lines = _ReadInformeParagraphs $file $wordApp } catch { $lines = @() }

    $gia = _ExtractIdGia $lines
    $exp = _ExtractExpedient $lines
    $font = 'document'
    if ([string]::IsNullOrWhiteSpace($gia)) {
        $gia = _GiaFromFolderName $file.FullName
        if (-not [string]::IsNullOrWhiteSpace($gia)) { $font = 'carpeta' }
    }
    if ([string]::IsNullOrWhiteSpace($gia) -and $null -ne $expToGia) {
        $key = _NormalitzaExpedient $exp
        if ($key -ne '' -and $expToGia.ContainsKey($key)) { $gia = $expToGia[$key]; $font = 'excel' }
    }

    $conclInfo = _ExtractConclusio $lines
    $concl = $conclInfo.Text

    $motius = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($gia)) { [void]$motius.Add('sense ID GIA') }
    $conclMotiu = _ConclusioMotiu $conclInfo
    if (-not [string]::IsNullOrWhiteSpace($conclMotiu)) { [void]$motius.Add($conclMotiu) }

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
        Conclusio     = $concl
        ConclusioBreu = (_ConclusioBreu $concl)
        Fitxer        = $file.Name
        Ruta          = $file.FullName
        Carpeta       = _CarpetaActivitat $file.FullName
        Modificat     = $file.LastWriteTimeUtc.ToString('o')
        Ignorat       = (_ConclusioIgnorarPerDefecte $conclInfo)
        Motius        = $motius.ToArray()
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
# plans indexats per 'ruta', arrossegant les dades de l'activitat a cada informe.
# Cada registre te la mateixa forma que Get-InformeData (perque el reagrupament
# els tracti igual). Retorna una hashtable [ruta] -> registre.
function _FlattenInformesDb($db) {
    $map = @{}
    if ($null -eq $db -or $null -eq $db.activitats) { return $map }
    foreach ($act in $db.activitats) {
        if ($null -eq $act.informes) { continue }
        foreach ($inf in $act.informes) {
            $ruta = [string]$inf.ruta
            if ([string]::IsNullOrWhiteSpace($ruta)) { continue }
            $motiuStr = if ($null -ne $inf.PSObject.Properties['motiu']) { [string]$inf.motiu } else { '' }
            $motius = if ([string]::IsNullOrWhiteSpace($motiuStr)) { @() } else { @($motiuStr -split ',\s*') }
            $ign = $false
            if ($null -ne $inf.PSObject.Properties['ignorat']) { $ign = [bool]$inf.ignorat }
            $conclusioText = [string]$inf.conclusio
            # Compatibilitat: si la base es d'abans d'aquest camp, la calculem
            # ara mateix (no cal reescanejar per tenir-la la primera vegada).
            $conclusioBreu = if ($null -ne $inf.PSObject.Properties['conclusio_breu']) { [string]$inf.conclusio_breu } else { _ConclusioBreu $conclusioText }
            $map[$ruta] = [pscustomobject]@{
                Data          = [string]$inf.data
                Gia           = [string]$act.id_gia
                GiaFont       = ''
                Expedient     = [string]$act.expedient
                Titular       = [string]$act.titular
                Conclusio     = $conclusioText
                ConclusioBreu = $conclusioBreu
                Fitxer        = [string]$inf.fitxer
                Ruta          = $ruta
                Carpeta       = [string]$act.carpeta
                Modificat     = if ($null -ne $inf.PSObject.Properties['modificat']) { [string]$inf.modificat } else { '' }
                Ignorat       = $ign
                Motius        = $motius
            }
        }
    }
    return $map
}

# ----------------------------------------------------------------------------
# Escaneig complet + escriptura del JSON (interactiu, amb finestra de progres)
# ----------------------------------------------------------------------------
function Invoke-InformesDbScan {
    # 1. Resoldre la carpeta d'informes. -ErrorAction SilentlyContinue: si la
    #    unitat (p.ex. la I: de la feina) no existeix, Test-Path no ha de petar,
    #    nomes ha de donar 'no trobada' (potser estas fora de la feina).
    $dir = $InformesDir
    $existeix = $false
    if (-not [string]::IsNullOrWhiteSpace($dir)) {
        try { $existeix = Test-Path -LiteralPath $dir -ErrorAction SilentlyContinue } catch { $existeix = $false }
    }
    if (-not $existeix) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat la carpeta d'informes:`n$dir`n`nSi treballes fora de la feina (sense la unitat I:), obre-la quan hi tinguis accés. Pots canviar la ruta amb `$InformesDir a config.ps1.",
            'Base d''informes', 'OK', 'Warning') | Out-Null
        return
    }

    # 2. Finestra de progres.
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

    try {
        # 3. Carregar l'Excel d'activitats (opcional; per la cerca inversa i el
        #    titular). Si no hi ha Excel, es continua sense aquest fallback.
        $cache = $null; $expToGia = $null
        try {
            $excel = Find-LatestActivitatsExcel
            if ($null -ne $excel) {
                $lbl.Text = "Llegint la base d'activitats (Excel)..."
                [System.Windows.Forms.Application]::DoEvents()
                $cache = Initialize-ActivitatsCache $excel.File
                $expToGia = Build-ExpedientToGiaMap $cache
            }
        } catch { $cache = $null; $expToGia = $null }

        # 3b. Carregar la base anterior (si existeix) per fer un escaneig
        #     INCREMENTAL: nomes es reobren els .docx modificats DESPRES de
        #     l'ultima actualitzacio; la resta es reutilitzen (conservant el seu
        #     "ignorat"). Els fitxers que ja no existeixen es podaran sols (nomes
        #     reagrupem els que trobem ara). Si no hi ha base previa (o esta
        #     corrupta), es fa un escaneig complet.
        $outPath    = Join-Path $LocalActivitatsDir 'informes-db.json'
        $prevByRuta = @{}
        $prevUtc    = [datetime]::MinValue
        $generatEl  = (Get-Date).ToString('o')
        if (Test-Path -LiteralPath $outPath) {
            try {
                $prevDb     = (Get-Content -LiteralPath $outPath -Raw -Encoding UTF8 | ConvertFrom-Json)
                $prevByRuta = _FlattenInformesDb $prevDb
                if ($prevDb.PSObject.Properties['actualitzat_el'] -and -not [string]::IsNullOrWhiteSpace([string]$prevDb.actualitzat_el)) {
                    try { $prevUtc = ([datetime]::Parse([string]$prevDb.actualitzat_el)).ToUniversalTime() } catch { $prevUtc = [datetime]::MinValue }
                }
                if ($prevDb.PSObject.Properties['generat_el'] -and -not [string]::IsNullOrWhiteSpace([string]$prevDb.generat_el)) {
                    $generatEl = [string]$prevDb.generat_el
                }
            } catch { $prevByRuta = @{}; $prevUtc = [datetime]::MinValue }
        }

        # 4. Recollir els fitxers candidats (.docx o .doc amb data al principi
        #    del nom). Un sol Get-ChildItem recursiu (sense -Filter) i filtrem
        #    per extensio nosaltres: evita el parany de "*.doc" -Filter que a
        #    vegades tambe encerta ".docx" pel nom curt (8.3) de NTFS.
        $lbl.Text = "Cercant informes a:`n$dir"
        [System.Windows.Forms.Application]::DoEvents()
        $allInformes = Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue |
                       Where-Object {
                           $_.Name -notlike '~$*' -and
                           ($_.Extension -ieq '.docx' -or $_.Extension -ieq '.doc') -and
                           $null -ne (_ParseDataInformeFromName $_.Name)
                       }
        $files = @($allInformes)
        $total = $files.Count

        $bar.Style = 'Continuous'
        $bar.Minimum = 0
        $bar.Maximum = [Math]::Max(1, $total)

        # 5. Analitzar cada informe (incremental: reutilitzem els no modificats).
        #    Word només es crea (mandrosament) si cal reprocessar algun .doc
        #    antic; es tanca sempre al 'finally', encara que hi hagi un error.
        $informes = New-Object System.Collections.ArrayList
        $revisar  = New-Object System.Collections.ArrayList
        $reprocessats = 0
        $i = 0
        $wordApp = $null
        try {
            foreach ($f in $files) {
                $i++
                $ruta = $f.FullName
                $teEntrada = $prevByRuta.ContainsKey($ruta)
                if (-not (_HaDeReprocessar $f.LastWriteTimeUtc $prevUtc $teEntrada)) {
                    # No s'ha tocat des de l'ultim escaneig: reutilitzem l'entrada.
                    $r = $prevByRuta[$ruta]
                } else {
                    if ($f.Extension -ieq '.doc' -and $null -eq $wordApp) {
                        try { $wordApp = New-Object -ComObject Word.Application; $wordApp.Visible = $false } catch { $wordApp = $null }
                    }
                    $r = Get-InformeData $f $expToGia $cache $wordApp
                    # Conservem l'"ignorat" i la "conclusio breu" que l'usuari
                    # hagi marcat/corregit abans (encara que l'informe s'hagi
                    # hagut de reprocessar).
                    if ($teEntrada) {
                        $r.Ignorat = [bool]$prevByRuta[$ruta].Ignorat
                        $r.ConclusioBreu = [string]$prevByRuta[$ruta].ConclusioBreu
                    }
                    $reprocessats++
                }
                if (($i % 5) -eq 0 -or $i -eq $total) {
                    $lbl.Text = "Analitzant informes... ($i de $total, $reprocessats de nous/modificats)"
                    $bar.Value = [Math]::Min($bar.Maximum, $i)
                    [System.Windows.Forms.Application]::DoEvents()
                }
                [void]$informes.Add($r)
                if ($r.Motius.Count -gt 0) {
                    [void]$revisar.Add([pscustomobject]@{
                        fitxer = $r.Fitxer
                        ruta   = $r.Ruta
                        motiu  = ($r.Motius -join ', ')
                    })
                }
            }
        } finally {
            if ($null -ne $wordApp) { try { $wordApp.Quit() } catch { } }
        }

        # 6. Agrupar per activitat: per ID GIA quan n'hi ha; si NO en tenen, per
        #    CARPETA (tots els informes d'una mateixa carpeta = una activitat).
        #    Ordenem els informes de cada activitat per data.
        $groups = [ordered]@{}
        foreach ($r in $informes) {
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
            if ([string]::IsNullOrWhiteSpace($g.id_gia)    -and -not [string]::IsNullOrWhiteSpace($r.Gia))       { $g.id_gia = $r.Gia }
            if ([string]::IsNullOrWhiteSpace($g.expedient) -and -not [string]::IsNullOrWhiteSpace($r.Expedient)) { $g.expedient = $r.Expedient }
            if ([string]::IsNullOrWhiteSpace($g.titular)   -and -not [string]::IsNullOrWhiteSpace($r.Titular))   { $g.titular = $r.Titular }
            [void]$g._informes.Add([pscustomobject]@{
                data          = $r.Data
                fitxer        = $r.Fitxer
                ruta          = $r.Ruta
                conclusio     = $r.Conclusio
                conclusio_breu = $r.ConclusioBreu
                modificat     = $r.Modificat
                ignorat       = [bool]$r.Ignorat
                motiu         = ($r.Motius -join ', ')
            })
        }

        $activitats = New-Object System.Collections.ArrayList
        foreach ($g in $groups.Values) {
            $ordered = @($g._informes | Sort-Object { if ($_.data) { $_.data } else { '' } })
            [void]$activitats.Add([pscustomobject]@{
                id_gia       = $g.id_gia
                expedient    = $g.expedient
                titular      = $g.titular
                carpeta      = $g.carpeta
                estat_actual = (_EstatActualActivitat $ordered)
                informes     = $ordered
            })
        }
        $activitatsOrd = @($activitats | Sort-Object { [string]$_.id_gia })

        # 7. Escriure el JSON (conservem generat_el; actualitzat_el = ara).
        $outObj = [pscustomobject]@{
            generat_el     = $generatEl
            actualitzat_el = (Get-Date).ToString('o')
            carpeta_arrel  = $dir
            n_informes     = $informes.Count
            n_activitats   = $activitatsOrd.Count
            activitats     = $activitatsOrd
            a_revisar      = @($revisar)
        }
        if (-not (Test-Path -LiteralPath $LocalActivitatsDir)) {
            New-Item -ItemType Directory -Path $LocalActivitatsDir -Force | Out-Null
        }
        ($outObj | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $outPath -Encoding UTF8

        $form.Close()

        # 8. Resum.
        $msg = "Base d'informes actualitzada.`n`n" +
               "Informes trobats: $($informes.Count)`n" +
               "Nous o modificats (reprocessats): $reprocessats`n" +
               "Activitats: $($activitatsOrd.Count)`n" +
               "A revisar: $($revisar.Count)`n`n" +
               "Fitxer:`n$outPath`n`nVols obrir-lo?"
        $r = [System.Windows.Forms.MessageBox]::Show($msg, 'Base d''informes', 'YesNo', 'Information')
        if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
            try { Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$outPath`"" | Out-Null } catch { }
        }
    }
    catch {
        try { $form.Close() } catch { }
        [System.Windows.Forms.MessageBox]::Show(
            "Error escanejant els informes:`n$($_.Exception.Message)",
            'Base d''informes', 'OK', 'Error') | Out-Null
    }
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
        $db = (Get-Content -LiteralPath $outPath -Raw -Encoding UTF8 | ConvertFrom-Json)
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

    # Agrupem visualment les files: primer les que tenen ID GIA (ordenades per
    # GIA), despres les que NO en tenen agrupades per CARPETA. Aixi, per als
    # informes sense GIA, els d'una mateixa carpeta queden JUNTS. Dins de cada
    # grup, per data.
    $allRows = @($allRows | Sort-Object `
        @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.Gia)) { 1 } else { 0 } } }, `
        @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.Gia)) { $_.Carpeta } else { $_.Gia } } }, `
        @{ Expression = { [string]$_.Data } })

    # Estat compartit amb els gestors d'esdeveniments (hashtable per referencia).
    # 'Loading' evita que el gestor de la casella reaccioni mentre s'omple la
    # graella (Rows.Add pot disparar CellValueChanged abans d'assignar el Tag).
    # SortColIdx/SortAsc: columna d'ordenacio SECUNDARIA triada per l'usuari
    # (-1 = cap). L'agrupament per activitat (GIA/carpeta) sempre es la clau
    # PRIMARIA i la data la darrera; la columna triada nomes ordena DINS de
    # cada activitat. Aixi mai es trenca l'agrupament, ordenis el que ordenis.
    $state = @{ Dirty = $false; Db = $db; Path = $outPath; Loading = $false; SortColIdx = -1; SortAsc = $true }

    $form = _NewForm
    $form.Text = "Editar base d'informes"
    $form.Size = New-Object System.Drawing.Size(1000, 676)
    $form.MinimumSize = New-Object System.Drawing.Size(720, 476)

    # Graella.
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AutoSizeColumnsMode = 'None'
    $grid.MultiSelect = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White

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
    # (mes avall) per mantenir sempre l'agrupament per activitat. Desactivem
    # l'ordenacio automatica de totes les columnes menys el boto "Obrir".
    foreach ($col in $grid.Columns) {
        if ($col.Index -ne $idxObrir) { $col.SortMode = 'Programmatic' }
    }

    $fontNormal = $grid.Font
    $fontStrike = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Strikeout)

    # Expressions d'ordenacio per index de columna (sobre la fila $_). El boto
    # "Obrir" (8) no s'ordena.
    $colExpr = @{
        0 = { [string]$_.Data }
        1 = { [string]$_.Gia }
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
    $lblCerca = New-Object System.Windows.Forms.Label
    $lblCerca.Text = 'Cerca:'; $lblCerca.AutoSize = $true
    $lblCerca.Location = New-Object System.Drawing.Point(10, 13)
    $txtCerca = New-Object System.Windows.Forms.TextBox
    $txtCerca.Location = New-Object System.Drawing.Point(62, 10); $txtCerca.Width = 300
    $topPanel.Controls.Add($lblCerca)
    $topPanel.Controls.Add($txtCerca)

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
            if ($n -ne '') {
                $hay = (($row.Data + ' ' + $row.Gia + ' ' + $row.Titular + ' ' + $row.Carpeta + ' ' + $row.Conclusio + ' ' + $row.ConclusioBreu + ' ' + $row.EstatActual + ' ' + $row.Motiu)).ToLower()
                if (-not $hay.Contains($n)) { continue }
            }
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

        # 1) agrupament per activitat (GIA, si no carpeta) SEMPRE primer;
        # 2) columna triada per l'usuari (asc/desc), nomes desempata dins
        #    l'activitat; 3) data com a darrera clau.
        $crit = @(
            @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.Gia)) { 1 } else { 0 } } },
            @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.Gia)) { $_.Carpeta } else { $_.Gia } } }
        )
        if ($state.SortColIdx -ge 0 -and $colExpr.ContainsKey($state.SortColIdx)) {
            $crit += @{ Expression = $colExpr[$state.SortColIdx]; Descending = (-not $state.SortAsc) }
        }
        $crit += @{ Expression = { [string]$_.Data } }
        $rows = @($rows | Sort-Object -Property $crit)

        foreach ($row in $rows) {
            $ign = [bool]$row.Obj.ignorat
            $idx = $grid.Rows.Add(@($row.Data, $row.Gia, $row.Titular, $row.Carpeta, $row.Conclusio, $row.ConclusioBreu, $row.EstatActual, $row.Motiu, 'Obrir', $ign))
            $gr = $grid.Rows[$idx]
            $gr.Tag = $row
            $gr.Cells[4].ToolTipText = $row.Conclusio
            _StyleInformeRow $gr $ign $fontNormal $fontStrike
        }
        $state.Loading = $false
    }.GetNewClosure()

    # Cerca i filtres tornen a omplir la graella.
    $txtCerca.add_TextChanged({ & $fill }.GetNewClosure())
    $cbBreu.add_SelectedIndexChanged({ & $fill }.GetNewClosure())
    $cbEstat.add_SelectedIndexChanged({ & $fill }.GetNewClosure())
    $cbMotiu.add_SelectedIndexChanged({ & $fill }.GetNewClosure())
    $cbIgn.add_SelectedIndexChanged({ & $fill }.GetNewClosure())

    # Clic a la capcalera: tria la columna d'ordenacio SECUNDARIA (dins de cada
    # activitat) i alterna asc/desc. L'agrupament per activitat no es trenca mai.
    $grid.add_ColumnHeaderMouseClick({
        param($s, $e)
        if ($e.ColumnIndex -lt 0 -or $e.ColumnIndex -eq $idxObrir) { return }
        if ($state.SortColIdx -eq $e.ColumnIndex) { $state.SortAsc = (-not $state.SortAsc) }
        else { $state.SortColIdx = $e.ColumnIndex; $state.SortAsc = $true }
        foreach ($c in $grid.Columns) { $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None }
        $grid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection =
            if ($state.SortAsc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending }
        & $fill
    }.GetNewClosure())

    # Desa la base (retorna $true si va be).
    $doSave = {
        try {
            ($state.Db | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $state.Path -Encoding UTF8
            $state.Dirty = $false
            return $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n$($_.Exception.Message)", 'Editar base d''informes', 'OK', 'Error') | Out-Null
            return $false
        }
    }.GetNewClosure()

    # Barra inferior: botons.
    $botPanel = New-Object System.Windows.Forms.Panel
    $botPanel.Dock = 'Bottom'; $botPanel.Height = 48
    $btnDesar = New-Object System.Windows.Forms.Button
    $btnDesar.Text = 'Desar'; $btnDesar.Size = New-Object System.Drawing.Size(120, 30)
    $btnDesar.Location = New-Object System.Drawing.Point(10, 9)
    _StylePrimaryButton $btnDesar
    $btnDesar.add_Click({
        if (& $doSave) {
            [System.Windows.Forms.MessageBox]::Show('Canvis desats.', 'Editar base d''informes', 'OK', 'Information') | Out-Null
        }
    }.GetNewClosure())
    $btnTancar = New-Object System.Windows.Forms.Button
    $btnTancar.Text = ([char]0x2190 + ' Enrere'); $btnTancar.Size = New-Object System.Drawing.Size(120, 30)
    $btnTancar.Location = New-Object System.Drawing.Point(140, 9)
    _StyleSecondaryButton $btnTancar
    $btnTancar.add_Click({ $form.Close() }.GetNewClosure())

    # Exportar a CSV els llistats d'activitats en Estat Requeriment i Precinte /
    # Cessament (usa l'estat en memoria, que ja reflecteix els canvis no desats).
    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'Exportar llistats (CSV)'; $btnExport.Size = New-Object System.Drawing.Size(190, 30)
    $btnExport.Location = New-Object System.Drawing.Point(275, 9)
    _StyleSecondaryButton $btnExport
    $btnExport.add_Click({ Export-EstatsActivitats $state.Db }.GetNewClosure())

    $botPanel.Controls.Add($btnDesar)
    $botPanel.Controls.Add($btnTancar)
    $botPanel.Controls.Add($btnExport)

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
        $nouEstat = _EstatActualActivitat $row.Act.informes
        $row.Act.estat_actual = $nouEstat
        foreach ($gr2 in $grid.Rows) {
            $row2 = $gr2.Tag
            if ($null -ne $row2 -and $row2.Act -eq $row.Act) {
                $row2.EstatActual = $nouEstat
                $gr2.Cells[$idxEstat].Value = $nouEstat
            }
        }
        foreach ($row3 in $allRows) {
            if ($row3.Act -eq $row.Act) { $row3.EstatActual = $nouEstat }
        }
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

# ============================================================================
# Copiar informes — eina INFORMES "Copiar informes"
# ============================================================================
# Còpia PLANA i incremental dels INFORMES Word de $InformesDir a
# $CopiaInformesDir:
#  - NOMÉS es consideren informes: *.doc / *.docx (ignora els temporals ~$...)
#    AMB DATA AL PRINCIPI DEL NOM (AAAA-MM-DD, AA_MM_DD, etc.) — mateix criteri
#    que "Actualitzar base" (_ParseDataInformeFromName). Qualsevol altre Word
#    (plantilles, esborranys, documents diversos) NO es copia;
#  - guarda la data de l'última còpia (copia-informes-state.json) i només mira
#    els fitxers modificats DESPRÉS;
#  - si el fitxer ja és al destí (mateix nom), NO es torna a copiar;
#  - MAI esborra res del destí (còpia additiva);
#  - mostra una finestra de progrés amb botó CANCEL·LAR i confirma abans de
#    copiar (amb el nombre d'informes) perquè mai comenci "a cegues".
function Invoke-CopiarInformes {
    if ([string]::IsNullOrWhiteSpace($CopiaInformesDir)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha configurat cap carpeta de còpia.`n`nVes a  ⚙ Configuració  i indica 'Carpeta on copiar els informes'.",
            'Copiar informes', 'OK', 'Information') | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($InformesDir) -or -not (Test-Path -LiteralPath $InformesDir -ErrorAction SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat la carpeta d'informes:`n$InformesDir", 'Copiar informes', 'OK', 'Warning') | Out-Null
        return
    }
    if ([System.IO.Path]::GetFullPath($CopiaInformesDir).TrimEnd('\') -ieq [System.IO.Path]::GetFullPath($InformesDir).TrimEnd('\')) {
        [System.Windows.Forms.MessageBox]::Show("La carpeta de còpia no pot ser la mateixa que la carpeta d'informes.", 'Copiar informes', 'OK', 'Warning') | Out-Null
        return
    }
    try {
        if (-not (Test-Path -LiteralPath $CopiaInformesDir)) {
            New-Item -ItemType Directory -Path $CopiaInformesDir -Force -ErrorAction Stop | Out-Null
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut crear la carpeta destí:`n$($_.Exception.Message)", 'Copiar informes', 'OK', 'Error') | Out-Null
        return
    }

    # Estat incremental (mateix patró que actualitzat_el de l'escaneig).
    $stateFile = Join-Path $LocalActivitatsDir 'copia-informes-state.json'
    $prevUtc   = [datetime]::MinValue
    $generatEl = (Get-Date).ToString('o')
    if (Test-Path -LiteralPath $stateFile) {
        try {
            $st = (Get-Content -LiteralPath $stateFile -Raw -Encoding UTF8 | ConvertFrom-Json)
            if ($st.PSObject.Properties['generat_el']) { $generatEl = [string]$st.generat_el }
            $prevDesti = if ($st.PSObject.Properties['desti']) { [string]$st.desti } else { '' }
            # Només reaprofitem la data si el destí és el mateix; si ha canviat,
            # fem una còpia completa (prevUtc = MinValue).
            if ($prevDesti -eq $CopiaInformesDir -and $st.PSObject.Properties['copiat_el']) {
                try { $prevUtc = ([datetime]::Parse([string]$st.copiat_el)).ToUniversalTime() } catch { $prevUtc = [datetime]::MinValue }
            }
        } catch { $prevUtc = [datetime]::MinValue }
    }

    # ---- Finestra de progrés amb CANCEL·LAR --------------------------------
    # Running: mentre és cert, la X de la finestra es tracta com a "cancel·lar"
    # (no es tanca de debò fins al 'finally'), així el bucle no toca mai controls
    # ja destruïts.
    $cancel = @{ Flag = $false; Running = $true }
    $form = _NewForm
    $form.Text = 'Copiar informes'
    $form.Size = New-Object System.Drawing.Size(560, 190)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 18)
    $lbl.Size = New-Object System.Drawing.Size(510, 60)
    $lbl.Text = "Cercant informes a:`n$InformesDir"
    $form.Controls.Add($lbl)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 88)
    $bar.Size = New-Object System.Drawing.Size(510, 22)
    $bar.Style = 'Marquee'
    $form.Controls.Add($bar)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel·lar'
    $btnCancel.Size = New-Object System.Drawing.Size(120, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(410, 118)
    _StyleSecondaryButton $btnCancel
    $btnCancel.add_Click({ $cancel.Flag = $true }.GetNewClosure())
    $form.Controls.Add($btnCancel)
    $form.add_FormClosing({
        param($s, $e)
        if ($cancel.Running) { $cancel.Flag = $true; $e.Cancel = $true }  # X = cancel·lar; el tancament real el fa el 'finally'
    }.GetNewClosure())
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $copied = 0; $skipped = 0; $errors = 0
    $cancelled = $false
    $nothing = $false
    $trobats = 0
    try {
        # 1. Enumerar només els INFORMES (data al principi del nom). Streaming
        #    amb DoEvents periòdic perquè la finestra respongui i es pugui
        #    cancel·lar; 'break' atura el pipeline.
        $files = New-Object System.Collections.ArrayList
        $seen = 0
        Get-ChildItem -LiteralPath $InformesDir -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
            $seen++
            if (($seen % 200) -eq 0) {
                $lbl.Text = "Cercant informes...`n$seen fitxers explorats  ·  $($files.Count) informes trobats"
                [System.Windows.Forms.Application]::DoEvents()
                if ($cancel.Flag) { break }
            }
            if ($_.Name -notlike '~$*' -and
                ($_.Extension -ieq '.docx' -or $_.Extension -ieq '.doc') -and
                $null -ne (_ParseDataInformeFromName $_.Name)) {
                [void]$files.Add($_)
            }
        }
        $trobats = $files.Count

        if ($cancel.Flag) {
            $cancelled = $true
        } else {
            # 2. Quins cal copiar de veritat: modificats després de l'última còpia
            #    i que encara no siguin al destí (mai recopiem ni esborrem).
            $lbl.Text = "Preparant la llista d'informes a copiar..."
            [System.Windows.Forms.Application]::DoEvents()
            $toCopy = New-Object System.Collections.ArrayList
            foreach ($f in $files) {
                if ($f.LastWriteTimeUtc -le $prevUtc) { continue }
                $dest = [System.IO.Path]::Combine($CopiaInformesDir, $f.Name)
                if (Test-Path -LiteralPath $dest) { $skipped++; continue }
                [void]$toCopy.Add($f)
            }

            if ($toCopy.Count -eq 0) {
                $nothing = $true
            } else {
                # 3. Confirmació explícita ABANS de copiar res.
                $rc = [System.Windows.Forms.MessageBox]::Show(
                    "Es copiaran $($toCopy.Count) informes (de $trobats trobats; $skipped ja hi són).`n`nDestí: $CopiaInformesDir`n`nVols continuar?",
                    'Copiar informes', 'YesNo', 'Question')
                if ($rc -ne [System.Windows.Forms.DialogResult]::Yes) {
                    $cancelled = $true
                } else {
                    # 4. Còpia amb barra de progrés i cancel·lació.
                    $bar.Style = 'Continuous'; $bar.Minimum = 0; $bar.Maximum = [Math]::Max(1, $toCopy.Count); $bar.Value = 0
                    $done = 0
                    foreach ($f in $toCopy) {
                        if ($cancel.Flag) { $cancelled = $true; break }
                        $dest = [System.IO.Path]::Combine($CopiaInformesDir, $f.Name)
                        try { Copy-Item -LiteralPath $f.FullName -Destination $dest -ErrorAction Stop; $copied++ }
                        catch { $errors++ }
                        $done++
                        if ($bar.Value -lt $bar.Maximum) { $bar.Value = $done }
                        $lbl.Text = "Copiant informes...  $done de $($toCopy.Count)`nCopiats: $copied   Errors: $errors"
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                }
            }
        }
    } finally {
        $cancel.Running = $false   # permet el tancament real de la finestra
        try { $form.Close() } catch { }
    }

    if ($nothing) {
        [System.Windows.Forms.MessageBox]::Show(
            "No hi ha informes nous per copiar.`n`nInformes trobats: $trobats  (ja copiats: $skipped)`nDestí: $CopiaInformesDir",
            'Copiar informes', 'OK', 'Information') | Out-Null
        return
    }

    # 5. Desem l'estat NOMÉS si s'ha completat (si s'ha cancel·lat, deixem la
    #    data com estava perquè la propera vegada es tornin a comprovar els que
    #    faltaven; els ja copiats se saltaran igualment per existència).
    if (-not $cancelled) {
        $newState = [pscustomobject]@{
            generat_el = $generatEl
            copiat_el  = (Get-Date).ToString('o')
            desti      = $CopiaInformesDir
        }
        try {
            $dirState = [System.IO.Path]::GetDirectoryName($stateFile)
            if ($dirState -and -not (Test-Path -LiteralPath $dirState)) { New-Item -ItemType Directory -Path $dirState -Force | Out-Null }
            ($newState | ConvertTo-Json) | Set-Content -LiteralPath $stateFile -Encoding UTF8
        } catch { }
    }

    $titol = if ($cancelled) { 'Còpia cancel·lada' } else { 'Còpia completada' }
    [System.Windows.Forms.MessageBox]::Show(
        "$titol`n`nInformes copiats: $copied`nJa existents (omesos): $skipped`nErrors: $errors`n`nDestí: $CopiaInformesDir",
        'Copiar informes', 'OK', 'Information') | Out-Null
}

# ============================================================================
# Comprovar Excel — eina INFORMES "Comprovar Excel"
# ============================================================================
# Llegeix els "Camp Info N - Nom/Valor" de la fulla "Estès" i els indexa per GIA
# (columna 1). Retorna @{ Ok; Map = @{ gia -> @(@{Nom;Valor}) }; Error }.
function _ReadExcelCampInfoPerGia {
    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) { return @{ Ok = $false; Error = "No s'ha trobat cap Excel d'activitats." } }
    $excel = $null
    try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
    if ($null -eq $excel) { return @{ Ok = $false; Error = "No s'ha pogut iniciar Microsoft Excel." } }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($latest.File.FullName, 0, $true)   # ReadOnly
        try {
            $found = _FindEstesSheet $wb
            $sh = $found.Sheet
            if ($null -eq $sh) { return @{ Ok = $false; Error = "No s'ha trobat la fulla 'Estès' a l'Excel." } }
            $data = $sh.UsedRange.Value2
            if ($null -eq $data) { return @{ Ok = $true; Map = @{} } }
            $rows = $data.GetLength(0)
            $cols = $data.GetLength(1)
            $headers = @(); for ($c = 1; $c -le $cols; $c++) { $headers += [string]$data[1, $c] }
            $pairs = _FindCampInfoPairs $headers
            $map = @{}
            for ($r = 2; $r -le $rows; $r++) {
                $cell = $data[$r, 1]
                if ($null -eq $cell) { continue }
                $gia = if ($cell -is [double]) {
                    if ([math]::Floor($cell) -eq $cell) { [string][int]$cell } else { [string]$cell }
                } else { [string]$cell }
                $gia = $gia.Trim()
                if ($gia -eq '') { continue }
                $list = @()
                foreach ($p in $pairs) {
                    $nom = [string]$data[$r, $p.NomCol]
                    $val = [string]$data[$r, $p.ValorCol]
                    if (-not [string]::IsNullOrWhiteSpace($nom)) { $list += @{ Nom = $nom.Trim(); Valor = $val.Trim() } }
                }
                $map[$gia] = $list
            }
            return @{ Ok = $true; Map = $map }
        } finally {
            try { $wb.Close($false) } catch { }
        }
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    } finally {
        try { $excel.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    }
}

# Finestra modal senzilla amb un llistat de resultats (només lectura).
function _ShowResultatWindow($titol, $subtitol, $text) {
    $form = _NewForm
    $form.Text = $titol
    $form.Size = New-Object System.Drawing.Size(680, 560)
    $form.MinimumSize = New-Object System.Drawing.Size(480, 360)
    $form.StartPosition = 'CenterScreen'

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true
    $tb.ReadOnly = $true
    $tb.ScrollBars = 'Vertical'
    $tb.Dock = 'Fill'
    $tb.Font = New-Object System.Drawing.Font('Consolas', 9.5)
    $tb.BackColor = [System.Drawing.Color]::White
    $tb.Text = $text

    $bot = New-Object System.Windows.Forms.Panel
    $bot.Dock = 'Bottom'; $bot.Height = 46
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = 'Tancar'; $btn.Size = New-Object System.Drawing.Size(110, 30)
    $btn.Location = New-Object System.Drawing.Point(10, 8)
    _StyleSecondaryButton $btn
    $btn.add_Click({ $form.Close() }.GetNewClosure())
    $bot.Controls.Add($btn)

    $form.Controls.Add($tb)
    $form.Controls.Add($bot)
    [void](_AddBrandHeader $form $titol $subtitol 56)
    [void]$form.ShowDialog()
}

# Desa (o actualitza) una marca de temps d'"última execució" a un JSON petit,
# conservant les altres propietats que ja hi hagi. S'usa perquè el menú pugui
# mostrar quan es va executar cada eina per últim cop.
function _SaveRunTimestamp($path, [string]$prop) {
    try {
        $dir = [System.IO.Path]::GetDirectoryName($path)
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $data = [ordered]@{}
        if (Test-Path -LiteralPath $path) {
            try {
                $ex = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
                foreach ($pr in $ex.PSObject.Properties) { $data[$pr.Name] = $pr.Value }
            } catch { }
        }
        $data[$prop] = (Get-Date).ToString('o')
        ([pscustomobject]$data | ConvertTo-Json) | Set-Content -LiteralPath $path -Encoding UTF8
    } catch { }
}

# Comprova que les activitats en Estat "Precinte / Cessament" (base d'informes)
# tinguin a l'Excel el Camp Info corresponent amb valor "SI". Llista les que no.
function Invoke-ComprovarExcel {
    $comprovarState = Join-Path $LocalActivitatsDir 'comprovar-excel-state.json'
    $outPath = Join-Path $LocalActivitatsDir 'informes-db.json'
    if (-not (Test-Path -LiteralPath $outPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Encara no hi ha cap base d'informes.`n`nExecuta primer 'Actualitzar base'.",
            'Comprovar Excel', 'OK', 'Information') | Out-Null
        return
    }
    try {
        $db = (Get-Content -LiteralPath $outPath -Raw -Encoding UTF8 | ConvertFrom-Json)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir la base d'informes:`n$($_.Exception.Message)", 'Comprovar Excel', 'OK', 'Error') | Out-Null
        return
    }
    $targets = @($db.activitats | Where-Object { [string]$_.estat_actual -eq 'Precinte / Cessament' })
    if ($targets.Count -eq 0) {
        _SaveRunTimestamp $comprovarState 'comprovat_el'
        [System.Windows.Forms.MessageBox]::Show("No hi ha cap activitat en Estat 'Precinte / Cessament' a la base d'informes.", 'Comprovar Excel', 'OK', 'Information') | Out-Null
        return
    }

    $res = _ReadExcelCampInfoPerGia
    if (-not $res.Ok) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir l'Excel:`n$($res.Error)", 'Comprovar Excel', 'OK', 'Error') | Out-Null
        return
    }
    $map = $res.Map
    # La comprovació s'ha executat de debò: registrem la marca de temps.
    _SaveRunTimestamp $comprovarState 'comprovat_el'

    $desact = New-Object System.Collections.ArrayList
    $noTrob = New-Object System.Collections.ArrayList
    $senseGia = New-Object System.Collections.ArrayList
    foreach ($act in $targets) {
        $gia = [string]$act.id_gia
        if ([string]::IsNullOrWhiteSpace($gia)) {
            [void]$senseGia.Add("     - " + [string]$act.titular + "  (carpeta: " + [string]$act.carpeta + ")")
            continue
        }
        $etiqueta = "GIA " + $gia + "  -  " + [string]$act.titular
        if (-not $map.ContainsKey($gia)) { [void]$noTrob.Add("     - " + $etiqueta); continue }
        if (-not (_ExcelActivitatActualitzada $map[$gia])) { [void]$desact.Add("     - " + $etiqueta); continue }
    }

    if ($desact.Count -eq 0 -and $noTrob.Count -eq 0 -and $senseGia.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            ("L'Excel està al dia.`n`nTotes les {0} activitats en Estat 'Precinte / Cessament' tenen a l'Excel un Camp Info 'REQUERIT PER DECRET?' o 'PRECINTE?' amb valor SI." -f $targets.Count),
            'Comprovar Excel', 'OK', 'Information') | Out-Null
        return
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(("Activitats en Estat 'Precinte / Cessament' a la base d'informes: {0}" -f $targets.Count))
    [void]$sb.AppendLine("Criteri: a l'Excel han de tenir un Camp Info 'REQUERIT PER DECRET?' o 'PRECINTE?' amb valor que comenci per SI.")
    [void]$sb.AppendLine("")
    if ($desact.Count -gt 0) {
        [void]$sb.AppendLine(("DESACTUALITZADES a l'Excel (sense el Camp Info amb SI): {0}" -f $desact.Count))
        foreach ($l in $desact) { [void]$sb.AppendLine($l) }
        [void]$sb.AppendLine("")
    }
    if ($noTrob.Count -gt 0) {
        [void]$sb.AppendLine(("NO trobades a l'Excel (GIA inexistent a la fulla Estès): {0}" -f $noTrob.Count))
        foreach ($l in $noTrob) { [void]$sb.AppendLine($l) }
        [void]$sb.AppendLine("")
    }
    if ($senseGia.Count -gt 0) {
        [void]$sb.AppendLine(("NO verificables (activitat sense GIA a la base d'informes): {0}" -f $senseGia.Count))
        foreach ($l in $senseGia) { [void]$sb.AppendLine($l) }
    }

    _ShowResultatWindow "Comprovació de l'Excel" "Activitats precintades pendents d'actualitzar a l'Excel" ($sb.ToString())
}
